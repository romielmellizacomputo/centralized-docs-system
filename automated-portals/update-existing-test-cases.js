import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases'; // Fetch from "Boards Test Cases"
const SHEETS_TO_SKIP = [
  'ToC', 
  'Roster', 
  'Issues', 
  'HELP',
  'Metrics Comparison',
  'Test Case Portal',
  'Test Scenario Portal',
  'Scenario Extractor',
  'Case Extractor',
  `Feature Change Log`,
  `Logs`,
  `UTILS`,
  'TEMPLATE'
];
const MAX_URLS = 20;
const START_DATA_ROW = 3; // Skip first 2 rows (headers) when checking for data

// Enhanced rate limiting configuration
const RATE_LIMITS = {
  BETWEEN_SHEETS: 15000,      // 15 seconds between processing different sheets
  BETWEEN_URLS: 65000,        // 65 seconds between processing different URLs (critical!)
  BETWEEN_OPERATIONS: 2000,   // 2 seconds between API operations
  QUOTA_EXCEEDED_WAIT: 90000  // 90 seconds if quota is exceeded
};

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

// Utility function to sleep
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Utility function to format time
function formatTime(ms) {
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = seconds % 60;
  if (minutes > 0) {
    return `${minutes}m ${remainingSeconds}s`;
  }
  return `${seconds}s`;
}

// Enhanced countdown with progress
async function cooldownWithProgress(ms, message = 'Cooling down') {
  console.log(`⏳ ${message} for ${formatTime(ms)}...`);
  const interval = 5000; // Update every 5 seconds
  let elapsed = 0;
  
  while (elapsed < ms) {
    const remaining = ms - elapsed;
    const progress = (elapsed / ms) * 100;
    const barLength = 30;
    const filled = Math.floor(barLength * progress / 100);
    const bar = '█'.repeat(filled) + '░'.repeat(barLength - filled);
    
    process.stdout.write(`\r   [${bar}] ${progress.toFixed(0)}% - ${formatTime(remaining)} remaining`);
    
    const sleepTime = Math.min(interval, remaining);
    await sleep(sleepTime);
    elapsed += sleepTime;
  }
  
  console.log(`\r   ✅ Cooldown complete!${' '.repeat(50)}`);
}

// Fetch URLs from the D column of "Boards Test Cases"
async function fetchUrls(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!D3:D17`;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  const values = response.data.values || [];

  console.log(`   📥 Fetching URLs from ${range}...`);
  
  const urls = [];
  
  // Process URLs sequentially with rate limiting
  for (let index = 0; index < values.length; index++) {
    const row = values[index];
    const rowIndex = index + 3;
    const cellRange = `${SHEET_NAME}!D${rowIndex}`;

    try {
      const text = row[0] || null;
      if (!text) {
        console.log(`   ⏭️  No text found for cell D${rowIndex}`);
        continue;
      }

      const linkResponse = await sheets.spreadsheets.get({
        spreadsheetId: SHEET_ID,
        ranges: [cellRange],
        includeGridData: true,
        fields: 'sheets.data.rowData.values.hyperlink'
      });

      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);

      const hyperlink = linkResponse.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0]?.hyperlink;

      if (hyperlink) {
        urls.push({ url: hyperlink, rowIndex });
        console.log(`   ✅ Found URL in cell D${rowIndex}`);
        continue;
      }

      const urlRegex = /https?:\/\/\S+/;
      const match = text.match(urlRegex);
      const url = match ? match[0] : null;

      if (!url) {
        console.log(`   ⚠️  No URL found in text for cell D${rowIndex}`);
        continue;
      }

      urls.push({ url, rowIndex });
      console.log(`   ✅ Found URL in cell D${rowIndex}`);
    } catch (error) {
      console.error(`   ❌ Error processing URL for cell D${rowIndex}:`, error.message);
    }
  }

  return urls.filter(entry => entry && entry.url);
}

async function logData(auth, message) {
  const sheets = google.sheets({ version: 'v4', auth });
  const logCell = 'B1';
  const timestamp = new Date().toISOString().replace('T', ' ').substring(0, 19);
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!${logCell}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[`[${timestamp}] ${message}`]] }
  });
  console.log(`   📋 ${message}`);
  await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24', 'B27', 'C32', 'C11'];
  const ranges = cellRefs.map(ref => `${sheetTitle}!${ref}`);

  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges,
  });

  const data = {};
  res.data.valueRanges.forEach((range, index) => {
    const value = range.values?.[0]?.[0] || null;
    data[cellRefs[index]] = value;
  });

  if (!data['C24']) return null;

  await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);

  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = meta.data.sheets.find(s => s.properties.title === sheetTitle);
  const sheetId = sheet.properties.sheetId;
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;

  return {
    C24: data['C24'],
    C3: data['C3'],
    C4: data['C4'],
    B27: data['B27'],
    C5: data['C5'],
    C6: data['C6'],
    C7: data['C7'],
    C11: data['C11'],
    C32: data['C32'],
    C13: data['C13'],
    C14: data['C14'],
    C15: data['C15'],
    C18: data['C18'],
    C19: data['C19'],
    C20: data['C20'],
    C21: data['C21'],
    sheetUrl,
    sheetName: sheetTitle
  };
}

async function processUrl(url, auth, urlIndex, totalUrls) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`🔄 Processing URL ${urlIndex}/${totalUrls}`);
  console.log(`${'='.repeat(70)}`);
  
  const targetSpreadsheetId = url.match(/[-\w]{25,}/)[0];
  const sheets = google.sheets({ version: 'v4', auth });

  const meta = await sheets.spreadsheets.get({ spreadsheetId: targetSpreadsheetId });
  const sheetTitles = meta.data.sheets
    .map(s => s.properties.title)
    .filter(title => !SHEETS_TO_SKIP.includes(title));

  console.log(`   📊 Found ${sheetTitles.length} sheets to process`);

  // Step 1: Collect all data using composite key (C24+C3)
  const sourceDataMap = new Map(); // Map of "C24|C3" -> data
  const sourceByC24 = new Map(); // Map of C24 -> Set of C3 values
  
  for (let i = 0; i < sheetTitles.length; i++) {
    const sheetTitle = sheetTitles[i];
    console.log(`   🔄 [${i + 1}/${sheetTitles.length}] Collecting data from: ${sheetTitle}`);
    
    const data = await collectSheetData(auth, targetSpreadsheetId, sheetTitle);
    
    if (data !== null && Object.values(data).some(v => v !== null && v !== '')) {
      if (data.C3 && data.C24) {
        const compositeKey = `${data.C24}|${data.C3}`;
        sourceDataMap.set(compositeKey, data);
        
        if (!sourceByC24.has(data.C24)) {
          sourceByC24.set(data.C24, new Set());
        }
        sourceByC24.get(data.C24).add(data.C3);
        
        console.log(`      ✅ Valid data collected for C24: ${data.C24}, C3: ${data.C3}`);
      } else {
        console.log(`      ⏭️  Missing C24 or C3 identifier`);
      }
    } else {
      console.log(`      ⏭️  No valid data`);
    }
    
    if (i < sheetTitles.length - 1) {
      await cooldownWithProgress(RATE_LIMITS.BETWEEN_SHEETS, `Cooling down before next sheet`);
    }
  }

  console.log(`\n   📦 Collected ${sourceDataMap.size} valid data entries`);
  console.log(`   📊 Source has ${sourceByC24.size} unique C24 groups`);
  await logData(auth, `Collected ${sourceDataMap.size} entries from URL: ${url}`);

  // Step 2: Validate and sync
  await validateAndSyncData(auth, sourceDataMap, sourceByC24);

  console.log(`   ✅ URL processing complete`);
}

async function validateAndSyncData(auth, sourceDataMap, sourceByC24) {
  const targetSheetTitles = await getTargetSheetTitles(auth);
  
  console.log(`\n   🔄 Syncing with target sheets...`);
  
  const sheetsToProcess = targetSheetTitles.filter(title => !SHEETS_TO_SKIP.includes(title));
  const sheetsToSkip = targetSheetTitles.filter(title => SHEETS_TO_SKIP.includes(title));
  
  console.log(`   📋 Will process ${sheetsToProcess.length} sheets`);
  if (sheetsToSkip.length > 0) {
    console.log(`   ⏭️  Will skip ${sheetsToSkip.length} protected sheets`);
  }

  for (const sheetTitle of targetSheetTitles) {
    if (SHEETS_TO_SKIP.includes(sheetTitle)) {
      console.log(`\n   ⏭️  Skipping: ${sheetTitle}`);
      continue;
    }

    console.log(`\n   🔄 Processing: ${sheetTitle}`);

    const isAllTestCases = sheetTitle === "ALL TEST CASES";
    const validateCol1 = isAllTestCases ? 'C' : 'B';
    const validateCol2 = isAllTestCases ? 'D' : 'C';

    const c24Column = await getColumnValues(auth, sheetTitle, validateCol1);
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    
    const c3Column = await getColumnValues(auth, sheetTitle, validateCol2);
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);

    const targetCompositeKeys = new Set();
    const rowsToDelete = [];
    const rowsToInsert = [];
    let updateCount = 0;

    // Find empty rows
    for (let i = 0; i < Math.max(c24Column.length, c3Column.length); i++) {
      const c24Value = c24Column[i];
      const c3Value = c3Column[i];
      const rowIndex = i + 1;
      
      if (rowIndex < START_DATA_ROW) continue;
      
      const c24Empty = !c24Value || c24Value.trim() === '';
      const c3Empty = !c3Value || c3Value.trim() === '';
      
      if (c24Empty && c3Empty) {
        rowsToDelete.push({ rowIndex, c24Value: '(empty)', c3Value: '(empty)', reason: 'Empty' });
        console.log(`      🗑️  Row ${rowIndex} - Empty row`);
      }
    }

    // Find rows to delete within source C24 groups
    for (let i = 0; i < c3Column.length; i++) {
      const c3Value = c3Column[i];
      const c24Value = c24Column[i];
      const rowIndex = i + 1;

      const c24Empty = !c24Value || c24Value.trim() === '';
      const c3Empty = !c3Value || c3Value.trim() === '';
      if (c24Empty && c3Empty) continue;
      if (!c3Value || !c24Value) continue;
      if (c3Value === '%' || c3Value === '0%' || /^[0-9]+%?$/.test(c3Value)) {
        console.log(`      ⏭️  Row ${rowIndex} - Invalid C3 '${c3Value}'`);
        continue;
      }
      if (rowIndex < START_DATA_ROW) continue;
      
      const compositeKey = `${c24Value}|${c3Value}`;
      targetCompositeKeys.add(compositeKey);

      // CRITICAL: Only check if C24 exists in THIS source
      if (!sourceByC24.has(c24Value)) {
        console.log(`      ⏭️  Row ${rowIndex} - C24 '${c24Value}' not in this source`);
        continue;
      }
      
      const sourceC3Set = sourceByC24.get(c24Value);
      if (!sourceC3Set.has(c3Value)) {
        const alreadyMarked = rowsToDelete.some(r => r.rowIndex === rowIndex);
        if (!alreadyMarked) {
          rowsToDelete.push({ rowIndex, c24Value, c3Value, reason: `Not in ${c24Value} group` });
          console.log(`      🗑️  Row ${rowIndex} - C3 '${c3Value}' not in ${c24Value} group`);
        }
      }
    }

    // Update matching rows
    for (let i = 0; i < c3Column.length; i++) {
      const c3Value = c3Column[i];
      const c24Value = c24Column[i];
      const rowIndex = i + 1;

      if (!c3Value || !c24Value) continue;
      if (c3Value === '%' || c3Value === '0%' || /^[0-9]+%?$/.test(c3Value)) continue;
      if (rowIndex < START_DATA_ROW) continue;

      const compositeKey = `${c24Value}|${c3Value}`;
      
      if (sourceDataMap.has(compositeKey)) {
        const sourceData = sourceDataMap.get(compositeKey);
        console.log(`      ✏️  Row ${rowIndex} - Updating`);
        
        await clearRowData(auth, sheetTitle, rowIndex, isAllTestCases);
        await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
        
        await insertDataInRow(auth, sheetTitle, rowIndex, sourceData, isAllTestCases ? 'C' : 'B', isAllTestCases ? 'R' : 'Q');
        await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
        
        await logData(auth, `Updated row ${rowIndex} in '${sheetTitle}'`);
        updateCount++;
      }
    }

    // Find rows to insert
    for (const [compositeKey, sourceData] of sourceDataMap.entries()) {
      if (!targetCompositeKeys.has(compositeKey)) {
        rowsToInsert.push({ compositeKey, data: sourceData });
        console.log(`      ➕ Will insert C24 '${sourceData.C24}', C3 '${sourceData.C3}'`);
      }
    }

    // Delete rows
    if (rowsToDelete.length > 0) {
      console.log(`\n      🗑️  Deleting ${rowsToDelete.length} rows...`);
      rowsToDelete.sort((a, b) => b.rowIndex - a.rowIndex);
      
      for (const { rowIndex, reason } of rowsToDelete) {
        try {
          await deleteRow(auth, sheetTitle, rowIndex);
          await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
          await logData(auth, `Deleted row ${rowIndex} - ${reason}`);
          console.log(`      ✅ Deleted row ${rowIndex}`);
        } catch (error) {
          if (error.message.includes('protected')) {
            console.log(`      ⚠️  Row ${rowIndex} protected`);
          } else {
            console.log(`      ❌ Error deleting row ${rowIndex}`);
          }
        }
      }
    }

    // Insert rows
    if (rowsToInsert.length > 0) {
      console.log(`\n      ➕ Inserting ${rowsToInsert.length} rows...`);
      
      for (const { data } of rowsToInsert) {
        let lastC24Index = -1;
        for (let i = 0; i < c24Column.length; i++) {
          if (c24Column[i] === data.C24) lastC24Index = i + 1;
        }

        if (lastC24Index !== -1) {
          const newRowIndex = lastC24Index + 1;
          await insertRowWithFormat(auth, sheetTitle, lastC24Index);
          await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
          
          await insertDataInRow(auth, sheetTitle, newRowIndex, data, isAllTestCases ? 'C' : 'B', isAllTestCases ? 'R' : 'Q');
          await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
          
          await logData(auth, `Inserted row for ${data.C24}|${data.C3}`);
          console.log(`      ✅ Inserted at row ${newRowIndex}`);
          
          c24Column.splice(lastC24Index, 0, data.C24);
        }
      }
    }

    console.log(`\n      📊 Summary: Updated ${updateCount}, Deleted ${rowsToDelete.length}, Inserted ${rowsToInsert.length}`);
    await sleep(RATE_LIMITS.BETWEEN_SHEETS);
  }
}

async function deleteRow(auth, sheetTitle, rowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });
  const sheetId = await getSheetId(auth, sheetTitle);

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId: sheetId,
            dimension: 'ROWS',
            startIndex: rowIndex - 1,
            endIndex: rowIndex
          }
        }
      }]
    }
  });
}

async function insertRowWithFormat(auth, sheetTitle, sourceRowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [{
        insertDimension: {
          range: {
            sheetId: await getSheetId(auth, sheetTitle),
            dimension: 'ROWS',
            startIndex: sourceRowIndex,
            endIndex: sourceRowIndex + 1
          },
          inheritFromBefore: true
        }
      }]
    }
  });
}

async function getSheetId(auth, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetTitle);
  return sheet.properties.sheetId;
}

async function insertDataInRow(auth, sheetTitle, row, data, startCol, endCol) {
  const sheets = google.sheets({ version: 'v4', auth });
  const isAllTestCases = sheetTitle === "ALL TEST CASES";

  const values = [
    data.C24,
    data.C3,
    `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
    data.B27,
    data.C5,
    data.C6,
    data.C7,
    '',
    data.C11,
    data.C32,
    data.C15,
    data.C13,
    data.C14,
    data.C18,
    data.C19,
    data.C20,
    data.C21
  ];

  if (isAllTestCases) values.push(data.C21);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!${startCol}${row}:${endCol}${row}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] }
  });
}

async function clearRowData(auth, sheetTitle, row, isAllTestCases) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = isAllTestCases ? `D${row}:T${row}` : `C${row}:S${row}`;
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SHEET_ID,
    range: `${sheetTitle}!${range}`
  });
}

async function getTargetSheetTitles(auth) {
  const sheets = google.sheets({ version: 'v4', auth });
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return meta.data.sheets.map(s => s.properties.title);
}

async function getColumnValues(auth, sheetTitle, column) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${sheetTitle}!${column}:${column}`;
  const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
  return res.data.values?.map(row => row[0]) || [];
}

async function updateTestCasesInLibrary() {
  console.log('='.repeat(70));
  console.log('  📋 Boards Test Cases Updater - Enhanced 1:1 Sync');
  console.log('='.repeat(70));
  console.log(`⏰ Started at: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}\n`);
  
  const startTime = Date.now();
  const authClient = await auth.getClient();
  
  console.log(`📥 Fetching URLs from "${SHEET_NAME}"...`);
  const urlsWithIndices = await fetchUrls(authClient);

  if (!urlsWithIndices.length) {
    await logData(authClient, 'No URLs to process.');
    console.log('\n⚠️  No URLs found.');
    return;
  }

  console.log(`\n✅ Found ${urlsWithIndices.length} URL(s)\n`);
  await logData(authClient, `Found ${urlsWithIndices.length} URL(s) to process`);

  const uniqueUrls = new Set();
  let successCount = 0;
  let failCount = 0;
  let duplicateCount = 0;

  for (let i = 0; i < urlsWithIndices.length; i++) {
    const { url } = urlsWithIndices[i];
    
    if (uniqueUrls.has(url)) {
      duplicateCount++;
      await logData(authClient, `Duplicate URL: ${url}`);
      console.log(`⚠️  [${i + 1}/${urlsWithIndices.length}] Skipping duplicate`);
      continue;
    }

    uniqueUrls.add(url);
    await logData(authClient, `Processing URL: ${url}`);
    
    try {
      await processUrl(url, authClient, i + 1, urlsWithIndices.length);
      successCount++;
    } catch (error) {
      failCount++;
      
      if (error.message.includes('Quota exceeded') || error.code === 429) {
        await logData(authClient, `Quota exceeded, retrying...`);
        await cooldownWithProgress(RATE_LIMITS.QUOTA_EXCEEDED_WAIT, 'Quota exceeded');
        
        try {
          await processUrl(url, authClient, i + 1, urlsWithIndices.length);
          successCount++;
          failCount--;
        } catch (retryError) {
          await logData(authClient, `Error on retry: ${retryError.message}`);
        }
      } else {
        await logData(authClient, `Error: ${error.message}`);
      }
    }

    if (i < urlsWithIndices.length - 1) {
      await cooldownWithProgress(
        RATE_LIMITS.BETWEEN_URLS, 
        `Cooling down (${i + 1}/${urlsWithIndices.length} complete)`
      );
    }
  }

  const totalTime = Date.now() - startTime;
  
  console.log('\n' + '='.repeat(70));
  console.log('  📊 SUMMARY');
  console.log('='.repeat(70));
  console.log(`✅ Success: ${successCount}/${urlsWithIndices.length}`);
  if (duplicateCount > 0) console.log(`⏭️  Duplicates: ${duplicateCount}`);
  if (failCount > 0) console.log(`❌ Failed: ${failCount}`);
  console.log(`⏱️  Time: ${formatTime(totalTime)}`);
  console.log(`⏰ Finished: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}`);
  console.log('='.repeat(70));

  await logData(authClient, "Complete - 1:1 sync with composite key validation");
}

updateTestCasesInLibrary().catch(console.error);
