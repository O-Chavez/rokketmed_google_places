import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import axios from 'axios';
import xlsx from 'xlsx';
import dotenv from 'dotenv';
import levenshtein from 'fast-levenshtein';
import cliProgress from 'cli-progress';
import readline from 'readline';

dotenv.config();

// Create readline interface to ask for user input
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Ask the user which sheet to process
function askForSheetNumber() {
  return new Promise((resolve) => {
    rl.question(
      'What sheet number would you like to process (1-3)? ',
      (answer) => {
        const sheetNumber = parseInt(answer);
        if (sheetNumber >= 1 && sheetNumber <= 3) {
          resolve(sheetNumber);
        } else {
          console.log('Invalid input. Please enter a number between 1 and 3.');
          resolve(askForSheetNumber()); // Ask again if invalid input
        }
      }
    );
  });
}

let sheetNumber;

const apiKey = process.env.GOOGLE_PLACES_API_KEY;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let inputExcelPath;
let outputJsonPath;
let requestCounterPath;
let lastProcessedPath;

const bar = new cliProgress.SingleBar({
  format:
    'Progress [{bar}] {percentage}% | Sheet:{sheetNumber} | Status: {status}',
  barCompleteChar: '#',
  barIncompleteChar: '-',
  hideCursor: true,
});

let currentStatus = 'Starting...';

function updateStatus(status) {
  currentStatus = status;
  bar.update({ status, sheetNumber });
}

function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[sheetNumber];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

function delayWithVariance(baseDelay, variance) {
  const randomVariance = Math.random() * variance; // Random value between 0 and variance
  const totalDelay = baseDelay + randomVariance;
  return new Promise((resolve) => setTimeout(resolve, totalDelay));
}

function getRequestCounter() {
  if (fs.existsSync(requestCounterPath)) {
    return JSON.parse(fs.readFileSync(requestCounterPath, 'utf8'));
  }
  return { count: 0, lastReset: new Date().toDateString() };
}

async function checkAndUpdateRequestCount() {
  let counter = getRequestCounter();

  if (counter.lastReset !== new Date().toDateString()) {
    counter = { count: 0, lastReset: new Date().toDateString() };
  }

  if (counter.count >= 1000) {
    const now = new Date();
    const resumeTime = new Date(now);
    resumeTime.setHours(now.getHours() + 24); // Set 24 hours from the current time

    updateStatus(
      `Paused at ${now.toLocaleString()}. Daily API limit reached. Resuming at ${resumeTime.toLocaleString()}.`
    );

    // Delay the process for 24 hours
    const waitTime = resumeTime - now;

    await delayWithVariance(2000, 1000);
    return 0;
  }

  counter.count++;
  fs.writeFileSync(requestCounterPath, JSON.stringify(counter, null, 2));
  return counter.count;
}

function getLastProcessedRow() {
  if (fs.existsSync(lastProcessedPath)) {
    return JSON.parse(fs.readFileSync(lastProcessedPath, 'utf8')).lastRow;
  }
  return 0;
}

function updateLastProcessedRow(row) {
  fs.writeFileSync(
    lastProcessedPath,
    JSON.stringify({ lastRow: row }, null, 2)
  );
}

function normalizeString(str) {
  return str.trim().toLowerCase().replace(/\s+/g, ' ');
}

function calculateStringSimilarity(str1, str2) {
  const normalizedStr1 = normalizeString(str1);
  const normalizedStr2 = normalizeString(str2);
  const distance = levenshtein.get(normalizedStr1, normalizedStr2);
  const maxLength = Math.max(normalizedStr1.length, normalizedStr2.length);
  return 1 - distance / maxLength;
}

async function fetchWithRetry(url, retries = 3) {
  for (let i = 0; i < retries; i++) {
    try {
      const response = await axios.get(url);
      if (response.status === 200) return response.data;
    } catch (error) {
      console.warn(
        `API request failed (${i + 1}/${retries}): ${error.message}`
      );
      await delayWithVariance(3000, 1000);
    }
  }
  return null;
}

async function fetchGooglePlacesData(businessName, address) {
  const url = `https://maps.googleapis.com/maps/api/place/textsearch/json?query=${encodeURIComponent(
    businessName
  )}+${encodeURIComponent(
    address
  )}&fields=address_components,adr_address,aspects,business_status,formatted_address,formatted_phone_number,geometry,html_attributions,icon,icon_background_color,icon_mask_base_uri,international_phone_number,name,opening_hours,permanently_closed,photos,place_id,plus_code,price_level,rating,reviews,types,url,user_ratings_total,utc_offset,utc_offset_minutes,vicinity,website&key=${apiKey}`;

  for (let attempt = 0; attempt < 3; attempt++) {
    const data = await fetchWithRetry(url);
    if (!data) continue;

    if (data.status === 'OVER_QUERY_LIMIT') {
      updateStatus('Rate limit exceeded, waiting 5 minutes...');
      await delayWithVariance(300000, 1000);
      continue;
    }

    if (data.status === 'OK' && data.results.length > 0) {
      let bestMatch = data.results[0];
      let highestSimilarity = 0;

      data.results.forEach((result) => {
        const nameSimilarity = calculateStringSimilarity(
          businessName,
          result.name
        );
        const addressSimilarity = calculateStringSimilarity(
          address,
          result.formatted_address
        );
        const overallSimilarity = (nameSimilarity + addressSimilarity) / 2;

        if (overallSimilarity > highestSimilarity) {
          highestSimilarity = overallSimilarity;
          bestMatch = result;
        }
      });

      return bestMatch;
    } else {
      // Store the location that wasnt found
      await storeNotFoundLocations([{ businessName, address, sheetNumber }]);
    }

    return null;
  }
  return null;
}

async function appendToJsonFile(filePath, data) {
  let existingData = [];
  if (fs.existsSync(filePath)) {
    existingData = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  }
  existingData.push(data);
  fs.writeFileSync(filePath, JSON.stringify(existingData, null, 2));
}

// Function to store locations not found in Google Places API
async function storeNotFoundLocations(notFoundLocations) {
  const notFoundJsonPath = path.join(
    __dirname,
    './data/not_found_locations.json'
  );
  try {
    // If the file doesn't exist, create a new one. If it does, append data.
    const existingData = fs.existsSync(notFoundJsonPath)
      ? JSON.parse(fs.readFileSync(notFoundJsonPath))
      : [];

    // Append the new not-found locations to the existing data
    const updatedData = [...existingData, ...notFoundLocations];
    fs.writeFileSync(notFoundJsonPath, JSON.stringify(updatedData, null, 2));
  } catch (error) {
    console.error('Error saving not found locations:', error);
  }
}

async function processLocations() {
  const locations = readExcelFile(inputExcelPath);
  let lastProcessedRow = getLastProcessedRow();

  // Initialize progress bar
  bar.start(locations.length, lastProcessedRow, {
    sheetNumber,
    status: 'Starting...',
  });

  for (let i = lastProcessedRow; i < locations.length; i++) {
    const { 'Business Name': businessName, 'Street Address': address } =
      locations[i];

    if (businessName && address) {
      updateStatus(`Processing row ${i + 1}/${locations.length}`);
      if ((await checkAndUpdateRequestCount()) === 0) return;

      const placesData = await fetchGooglePlacesData(businessName, address);
      if (placesData) {
        await appendToJsonFile(outputJsonPath, placesData);
      }

      updateLastProcessedRow(i + 1);
      bar.update(i + 1);

      if (i < locations.length - 1) {
        await delayWithVariance(2500, 1000);
      }
    }
  }

  updateStatus(`Sheet:${sheetNumber} - Processing completed.`);
  bar.stop();
}

askForSheetNumber().then((number) => {
  sheetNumber = number;
  inputExcelPath = path.join(__dirname, './data/Rokketmed_Location_Data.xlsx');
  outputJsonPath = path.join(
    __dirname,
    `./data/output_sheet${sheetNumber}.json`
  );
  requestCounterPath = path.join(__dirname, './data/request_counter.json');
  lastProcessedPath = path.join(
    __dirname,
    `./data/last_processed_sheet${sheetNumber}.json`
  );

  processLocations().catch((error) => {
    console.error('Error processing locations:', error);
  });
});
