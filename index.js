import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
import axios from 'axios';
import xlsx from 'xlsx';
import dotenv from 'dotenv';
import levenshtein from 'fast-levenshtein';
import cliProgress from 'cli-progress';

dotenv.config();
const apiKey = process.env.GOOGLE_PLACES_API_KEY;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const inputExcelPath = path.join(
  __dirname,
  './data/Rokketmed_Location_Data.xlsx'
);
const outputJsonPath = path.join(__dirname, './data/output.json');
const requestCounterPath = path.join(__dirname, './data/request_counter.json');
const lastProcessedPath = path.join(__dirname, './data/last_processed.json');
const noLocationFound = path.join(__dirname, './data/no_location_found.json');

const bar = new cliProgress.SingleBar({
  format: 'Progress [{bar}] {percentage}% | Status: {status}',
  barCompleteChar: '#',
  barIncompleteChar: '-',
  hideCursor: true,
});

let currentStatus = 'Starting...';

function updateStatus(status) {
  currentStatus = status;
  bar.update({ status });
}

function readExcelFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  /*

  Sheets 1-3 (zero based) are the first sheet (index 0) location's broken into 3000 chunks.

  I've started on sheet 1.

  Adjust the 'workbook.SheetNames' index below to swap to next sheet

*/
  const sheetName = workbook.SheetNames[1];
  const sheet = workbook.Sheets[sheetName];
  return xlsx.utils.sheet_to_json(sheet);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
    updateStatus('Paused - Daily API limit reached');
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setHours(0, 0, 0, 0);
    const waitTime = tomorrow - now;
    await delay(waitTime);
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
      await delay(2000 * (i + 1));
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
      console.log('Rate limit exceeded, waiting 5 minutes...');
      await delay(5 * 60 * 1000);
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
    }

    // Handle no results found??

    // console.log(`No results found for ${businessName} at ${address}`);
    // const noLocationFound = path.join(__dirname, './data/no_location_found.json');

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

async function processLocations() {
  const locations = readExcelFile(inputExcelPath);
  let lastProcessedRow = getLastProcessedRow();

  // Initialize progress bar
  bar.start(locations.length, lastProcessedRow, { status: 'Starting...' });

  for (let i = lastProcessedRow; i < locations.length; i++) {
    const { 'Business Name': businessName, 'Street Address': address } =
      locations[i];

    if (businessName && address) {
      updateStatus(`Processing row ${i + 1}/${locations.length}`);
      if ((await checkAndUpdateRequestCount()) === 0) return;

      const placesData = await fetchGooglePlacesData(businessName, address);
      if (placesData) {
        await appendToJsonFile(outputJsonPath, placesData);
        updateLastProcessedRow(i + 1);
      }

      bar.update(i + 1);

      if (i < locations.length - 1) {
        await delay(1000);
      }
    }
  }

  updateStatus('Processing completed.');
  bar.stop();
  console.log('Processing completed.');
}

processLocations().catch((error) => {
  console.error('Error processing locations:', error);
});
