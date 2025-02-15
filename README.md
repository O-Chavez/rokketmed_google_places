## Overview

- This application is designed to run continuously and pause for 24 hours after the api limit is reached. Requests are manually throttled to 1/s to stay within api limitation.
- The original reference locations are stored in an excel file in the project at './data/Rokketmed_Location_data.xlsx'. This file contains 4 separate tables within - the original 9,000 locations along with three tables (sheets) breaking the 9,000 into 3,000 chunks.
- The process will need to be ran separately for each sheet/table and will continue until the current table is complete.
- Each sheets output will be saved to a separate json file noted as 'output_sheet' with the sheet number attached.
- Progress for each individual sheet is tracked in the project. If you need to stop the process, you can start it again without loosing progress.

#### Limitations

The Google Places has a free tier limit of 1,000 requests per day and 60 requests per minute. The status bar will state when a rate limit has been reach and will automatically restart the process once 24 hours has passed (if the 1,000 limit has been reached).

# Instructions

1. Run the following in the terminal after downloading and naviaged to the project in your console:

```console
npm i
```

2. Replace the API key in the .env file with your Google Places API key

3. Once all of the dependencies have been downloaded, you can begin the program by running:

```console
node index.js
```

4. Once the program beings, you will be asked which sheet you would like to begin processing. enter that number and the application will begin processing the locations.
