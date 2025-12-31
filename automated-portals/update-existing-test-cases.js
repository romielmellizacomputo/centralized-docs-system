import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases'; // Fetch from "Boards Test Cases"
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues', "HELP"];
const MAX_URLS = 20;

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

// Extract spreadsheet ID from URL
function extractSpreadsheetId(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
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

async function processUrl(spreadsheetId, auth, urlIndex, totalUrls) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`🔄 Processing Spreadsheet ${urlIndex}/${totalUrls}`);
  console.log(`   📄 Spreadsheet ID: ${spreadsheetId}`);
  console.log(`${'='.repeat(70)}`);
  
  const sheets = google.sheets({ version: 'v4', auth });

  const meta = await sheets.spreadsheets.get({ spreadsheetId: spreadsheetId });
  const sheetTitles = meta.data.sheets
    .map(s => s.properties.title)
    .filter(title => !SHEETS_TO_SKIP.includes(title));

  console.log(`   📊 Found ${sheetTitles.length} sheets to process`);

  // Process sheets SEQUENTIALLY with rate limiting (not parallel!)
  const validData = [];
  for (let i = 0; i < sheetTitles.length; i++) {
    const sheetTitle = sheetTitles[i];
    console.log(`   🔄 [${i + 1}/${sheetTitles.length}] Collecting data from: ${sheetTitle}`);
    
    const data = await collectSheetData(auth, spreadsheetId, sheetTitle);
    
    if (data !== null && Object.values(data).some(v => v !== null && v !== '')) {
      validData.push(data);
      console.log(`      ✅ Valid data collected`);
    } else {
      console.log(`      ⏭️  No valid data`);
    }
    
    // Rate limit between collecting data from different sheets
    if (i < sheetTitles.length - 1) {
      await cooldownWithProgress(RATE_LIMITS.BETWEEN_SHEETS, `Cooling down before next sheet`);
    }
  }

  console.log(`\n   📦 Processing ${validData.length} valid data entries...`);

  for (let i = 0; i < validData.length; i++) {
    const data = validData[i];
    console.log(`   🔄 [${i + 1}/${validData.length}] Validating and inserting: ${data.sheetName}`);
    await logData(auth, `Fetched data from sheet: ${data.sheetName}`);
    await validateAndInsertData(auth, data);
  }

  console.log(`   ✅ Spreadsheet processing complete`);
}

async function validateAndInsertData(auth, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  const targetSheetTitles = await getTargetSheetTitles(auth);
  let processed = false;

  for (const sheetTitle of targetSheetTitles) {
    if (SHEETS_TO_SKIP.includes(sheetTitle)) continue;

    const isAllTestCases = sheetTitle === "ALL TEST CASES";
    const validateCol1 = isAllTestCases ? 'C' : 'B';
    const validateCol2 = isAllTestCases ? 'D' : 'C';

    const firstColumn = await getColumnValues(auth, sheetTitle, validateCol1);
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    
    const secondColumn = await getColumnValues(auth, sheetTitle, validateCol2);
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);

    let lastC24Index = -1;
    let existingC3Index = -1;

    for (let i = 0; i < firstColumn.length; i++) {
      if (firstColumn[i] === data.C24) lastC24Index = i + 1;
      if (secondColumn[i] === data.C3) {
        existingC3Index = i + 1;
        break;
      }
    }

    if (existingC3Index !== -1) {
      await clearRowData(auth, sheetTitle, existingC3Index, isAllTestCases);
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      await insertDataInRow(auth, sheetTitle, existingC3Index, data, isAllTestCases ? 'C' : 'B');
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      await logData(auth, `Updated row ${existingC3Index} in sheet '${sheetTitle}'`);
      processed = true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await insertRowWithFormat(auth, sheetTitle, lastC24Index);
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      await insertDataInRow(auth, sheetTitle, newRowIndex, data, isAllTestCases ? 'C' : 'B');
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      await logData(auth, `Inserted row after ${lastC24Index} in sheet '${sheetTitle}'`);
      processed = true;
    }

    // Rate limiting between processing different target sheets
    await sleep(RATE_LIMITS.BETWEEN_SHEETS);
  }

  if (!processed) {
    await logData(auth, `No matches found for C24 ('${data.C24}') or C3 ('${data.C3}') in any sheet.`);
  }

  return processed;
}

async function insertRowWithFormat(auth, sheetTitle, sourceRowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: await getSheetId(auth, sheetTitle),
              dimension: 'ROWS',
              startIndex: sourceRowIndex,
              endIndex: sourceRowIndex + 1
            },
            inheritFromBefore: true
          }
        }
      ]
    }
  });

  console.log(`      📝 Inserted new row after row ${sourceRowIndex} in sheet '${sheetTitle}'`);
}

async function getSheetId(auth, sheetTitle) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetTitle);
  return sheet.properties.sheetId;
}

function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num *= 26;
    num += col.charCodeAt(i) - 64;
  }
  return num;
}

function numberToColumn(n) {
  let result = '';
  while (n > 0) {
    let remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

async function insertDataInRow(auth, sheetTitle, row, data, startCol) {
  const sheets = google.sheets({ version: 'v4', auth });

  const isAllTestCases = sheetTitle === "ALL TEST CASES";

  const rowValues = [
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

  if (isAllTestCases) {
    rowValues.push(data.C21);
  }

  const startIndex = columnToNumber(startCol);
  const endIndex = startIndex + rowValues.length - 1;
  const endCol = numberToColumn(endIndex);

  const range = `${sheetTitle}!${startCol}${row}:${endCol}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [rowValues]
    }
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
  console.log('  📋 Boards Test Cases Updater with Rate Limiting');
  console.log('='.repeat(70));
  console.log(`⏰ Started at: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}\n`);
  
  const startTime = Date.now();
  const authClient = await auth.getClient();
  
  console.log(`📥 Fetching URLs from sheet "${SHEET_NAME}"...`);
  const urlsWithIndices = await fetchUrls(authClient);

  if (!urlsWithIndices.length) {
    await logData(authClient, 'No URLs to process.');
    console.log('\n⚠️  No URLs found to process.');
    return;
  }

  console.log(`\n✅ Found ${urlsWithIndices.length} URL(s) to process\n`);

  // Group URLs by spreadsheet ID and track source rows
  const spreadsheetMap = new Map();
  
  for (const { url, rowIndex } of urlsWithIndices) {
    const spreadsheetId = extractSpreadsheetId(url);
    
    if (!spreadsheetId) {
      console.log(`⚠️  Could not extract spreadsheet ID from URL in row D${rowIndex}`);
      await logData(authClient, `Invalid URL format in row D${rowIndex}: ${url}`);
      continue;
    }
    
    if (!spreadsheetMap.has(spreadsheetId)) {
      spreadsheetMap.set(spreadsheetId, {
        url: url,
        rows: [rowIndex],
        spreadsheetId: spreadsheetId
      });
    } else {
      spreadsheetMap.get(spreadsheetId).rows.push(rowIndex);
    }
  }

  const uniqueSpreadsheets = Array.from(spreadsheetMap.values());
  const totalUrls = urlsWithIndices.length;
  const duplicateCount = totalUrls - uniqueSpreadsheets.length;

  console.log(`📊 Analysis:`);
  console.log(`   Total URLs found: ${totalUrls}`);
  console.log(`   Unique spreadsheets: ${uniqueSpreadsheets.length}`);
  if (duplicateCount > 0) {
    console.log(`   Duplicate references: ${duplicateCount}`);
    console.log(`\n📝 Spreadsheet grouping:`);
    uniqueSpreadsheets.forEach((entry, index) => {
      if (entry.rows.length > 1) {
        console.log(`   ${index + 1}. Spreadsheet ${entry.spreadsheetId}`);
        console.log(`      Referenced in rows: D${entry.rows.join(', D')}`);
      }
    });
  }
  console.log();

  await logData(authClient, `Found ${totalUrls} URL(s), ${uniqueSpreadsheets.length} unique spreadsheet(s) to process`);

  let successCount = 0;
  let failCount = 0;

  for (let i = 0; i < uniqueSpreadsheets.length; i++) {
    const { spreadsheetId, url, rows } = uniqueSpreadsheets[i];
    
    const rowsText = rows.length > 1 
      ? `rows D${rows.join(', D')}` 
      : `row D${rows[0]}`;
    
    await logData(authClient, `Processing spreadsheet ${spreadsheetId} (from ${rowsText})`);
    
    try {
      await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length);
      successCount++;
    } catch (error) {
      failCount++;
      
      if (error.message.includes('Quota exceeded') || error.code === 429) {
        await logData(authClient, `⚠️  Quota exceeded for spreadsheet: ${spreadsheetId}. Retrying after cooldown...`);
        await cooldownWithProgress(RATE_LIMITS.QUOTA_EXCEEDED_WAIT, 'Quota exceeded - cooling down');
        
        try {
          await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length);
          successCount++;
          failCount--;
        } catch (retryError) {
          await logData(authClient, `❌ Error processing spreadsheet on retry: ${spreadsheetId}. Error: ${retryError.message}`);
        }
      } else {
        await logData(authClient, `❌ Error processing spreadsheet: ${spreadsheetId}. Error: ${error.message}`);
      }
    }

    // CRITICAL: Add cooldown between processing different spreadsheets
    if (i < uniqueSpreadsheets.length - 1) {
      await cooldownWithProgress(
        RATE_LIMITS.BETWEEN_URLS, 
        `Cooling down before next spreadsheet (${i + 1}/${uniqueSpreadsheets.length} complete)`
      );
    }
  }

  const totalTime = Date.now() - startTime;
  
  console.log('\n' + '='.repeat(70));
  console.log('  📊 PROCESSING SUMMARY');
  console.log('='.repeat(70));
  console.log(`📄 Total URLs found: ${totalUrls}`);
  console.log(`📚 Unique spreadsheets processed: ${uniqueSpreadsheets.length}`);
  if (duplicateCount > 0) {
    console.log(`🔗 Duplicate references detected: ${duplicateCount}`);
  }
  console.log(`✅ Successful: ${successCount}/${uniqueSpreadsheets.length}`);
  if (failCount > 0) {
    console.log(`❌ Failed: ${failCount}/${uniqueSpreadsheets.length}`);
  }
  console.log(`⏱️  Total time: ${formatTime(totalTime)}`);
  console.log(`⏰ Finished at: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}`);
  console.log('='.repeat(70));

  await logData(authClient, "Processing complete.");
}

updateTestCasesInLibrary().catch(console.error);
