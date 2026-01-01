import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues', "HELP"];
const MAX_RUNTIME = 5.5 * 60 * 60 * 1000; // 5.5 hours for safety

// Optimized rate limiting
const RATE_LIMITS = {
  BETWEEN_READS: 500,
  BETWEEN_SHEETS: 2000,
  BETWEEN_SPREADSHEETS: 3000,
  BEFORE_WRITE_PHASE: 5000,
  BETWEEN_WRITES: 1000,
  BETWEEN_BATCH_WRITES: 2000,
  QUOTA_RECOVERY: 120000
};

// Track API calls
let apiCallCount = 0;
let lastResetTime = Date.now();
const MAX_CALLS_PER_MINUTE = 45;
const MINUTE = 60000;

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function trackApiCall() {
  apiCallCount++;
  const now = Date.now();
  
  if (now - lastResetTime > MINUTE) {
    console.log(`   📊 API counter reset. Previous minute: ${apiCallCount} calls`);
    apiCallCount = 1;
    lastResetTime = now;
    return;
  }
  
  if (apiCallCount >= MAX_CALLS_PER_MINUTE - 5) {
    const waitTime = MINUTE - (now - lastResetTime) + 3000;
    console.log(`   ⚠️  Rate limit protection: pausing for ${formatTime(waitTime)}`);
    await cooldownWithProgress(waitTime, 'Rate limit protection');
    apiCallCount = 1;
    lastResetTime = Date.now();
  }
}

function formatTime(ms) {
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;
  const remainingSeconds = seconds % 60;
  
  if (hours > 0) return `${hours}h ${remainingMinutes}m ${remainingSeconds}s`;
  if (minutes > 0) return `${minutes}m ${remainingSeconds}s`;
  return `${seconds}s`;
}

function extractSpreadsheetId(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

async function cooldownWithProgress(ms, message = 'Cooling down') {
  console.log(`⏳ ${message} for ${formatTime(ms)}...`);
  const interval = 2000;
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
  
  console.log(`\r   ✅ Complete!${' '.repeat(50)}`);
}

async function withRetry(fn, maxRetries = 2, baseDelay = RATE_LIMITS.QUOTA_RECOVERY) {
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      return await fn();
    } catch (error) {
      const isQuotaError = error.message.includes('Quota exceeded') || error.code === 429;
      const isLastAttempt = attempt === maxRetries;
      
      if (isQuotaError && !isLastAttempt) {
        const delay = baseDelay * (attempt + 1);
        console.log(`   ⚠️  Quota error (attempt ${attempt + 1}/${maxRetries + 1}), waiting ${formatTime(delay)}`);
        await cooldownWithProgress(delay, `Quota recovery attempt ${attempt + 1}`);
        continue;
      }
      
      throw error;
    }
  }
}

async function fetchAllUrls(auth) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const range = `${SHEET_NAME}!D3:D`;
    
    console.log(`   🔍 Fetching all URLs from ${range}...`);
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
    await sleep(RATE_LIMITS.BETWEEN_READS);
    
    const values = response.data.values || [];
    console.log(`   📥 Retrieved ${values.length} total rows`);
    
    const urls = [];
    
    for (let index = 0; index < values.length; index++) {
      const row = values[index];
      const actualRowIndex = index + 3;
      const cellRange = `${SHEET_NAME}!D${actualRowIndex}`;

      try {
        const text = row[0] || null;
        if (!text) {
          continue;
        }

        await trackApiCall();
        const linkResponse = await sheets.spreadsheets.get({
          spreadsheetId: SHEET_ID,
          ranges: [cellRange],
          includeGridData: true,
          fields: 'sheets.data.rowData.values.hyperlink'
        });

        await sleep(RATE_LIMITS.BETWEEN_READS);

        const hyperlink = linkResponse.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0]?.hyperlink;

        if (hyperlink) {
          urls.push({ url: hyperlink, rowIndex: actualRowIndex });
          continue;
        }

        const urlRegex = /https?:\/\/\S+/;
        const match = text.match(urlRegex);
        const url = match ? match[0] : null;

        if (url) {
          urls.push({ url, rowIndex: actualRowIndex });
        }
      } catch (error) {
        console.error(`   ❌ Error at D${actualRowIndex}: ${error.message}`);
      }
    }

    return urls.filter(entry => entry && entry.url);
  });
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24', 'B27', 'C32', 'C11'];
    const ranges = cellRefs.map(ref => `${sheetTitle}!${ref}`);

    const res = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges });
    await sleep(RATE_LIMITS.BETWEEN_READS);

    const data = {};
    res.data.valueRanges.forEach((range, index) => {
      data[cellRefs[index]] = range.values?.[0]?.[0] || null;
    });

    if (!data['C24']) return null;

    await trackApiCall();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    await sleep(RATE_LIMITS.BETWEEN_READS);
    
    const sheet = meta.data.sheets.find(s => s.properties.title === sheetTitle);
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.properties.sheetId}`;

    return { ...data, sheetUrl, sheetName: sheetTitle };
  });
}

async function collectDataFromSpreadsheet(spreadsheetId, auth, urlIndex, totalUrls, startTime) {
  const elapsed = Date.now() - startTime;
  if (elapsed > MAX_RUNTIME) {
    console.log('⏰ Timeout limit approaching');
    return { timeout: true, data: [] };
  }
  
  console.log(`\n${'='.repeat(70)}`);
  console.log(`🔄 Spreadsheet ${urlIndex}/${totalUrls} - ID: ${spreadsheetId}`);
  console.log(`   ⏱️  Elapsed: ${formatTime(elapsed)} | API calls: ${apiCallCount}`);
  console.log(`${'='.repeat(70)}`);
  
  try {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    await sleep(RATE_LIMITS.BETWEEN_READS);
    
    const sheetTitles = meta.data.sheets
      .map(s => s.properties.title)
      .filter(title => !SHEETS_TO_SKIP.includes(title));

    console.log(`   📊 Found ${sheetTitles.length} sheets to process`);

    const collectedData = [];
    for (let i = 0; i < sheetTitles.length; i++) {
      console.log(`   🔄 [${i + 1}/${sheetTitles.length}] Collecting: ${sheetTitles[i]}`);
      
      const data = await collectSheetData(auth, spreadsheetId, sheetTitles[i]);
      
      if (data && Object.values(data).some(v => v !== null && v !== '')) {
        collectedData.push(data);
        console.log(`      ✅ Data collected`);
      } else {
        console.log(`      ⏭️  No valid data`);
      }
      
      if (i < sheetTitles.length - 1) {
        await sleep(RATE_LIMITS.BETWEEN_SHEETS);
      }
    }

    console.log(`   ✅ Collected ${collectedData.length} valid entries`);
    return { timeout: false, data: collectedData };
  } catch (error) {
    console.error(`   ❌ Error: ${error.message}`);
    return { timeout: false, data: [], error: error.message };
  }
}

async function getTargetSheetStructure(auth) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    await sleep(RATE_LIMITS.BETWEEN_READS);
    
    const sheetTitles = meta.data.sheets.map(s => s.properties.title);
    const structure = {};
    
    for (const sheetTitle of sheetTitles) {
      if (SHEETS_TO_SKIP.includes(sheetTitle)) continue;
      
      const isAllTestCases = sheetTitle === "ALL TEST CASES";
      const validateCol1 = isAllTestCases ? 'C' : 'B';
      const validateCol2 = isAllTestCases ? 'D' : 'C';
      
      await trackApiCall();
      const firstColumn = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${sheetTitle}!${validateCol1}:${validateCol1}`
      });
      await sleep(RATE_LIMITS.BETWEEN_READS);
      
      await trackApiCall();
      const secondColumn = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${sheetTitle}!${validateCol2}:${validateCol2}`
      });
      await sleep(RATE_LIMITS.BETWEEN_READS);
      
      structure[sheetTitle] = {
        isAllTestCases,
        firstColumn: firstColumn.data.values?.map(row => row[0]) || [],
        secondColumn: secondColumn.data.values?.map(row => row[0]) || [],
        sheetId: meta.data.sheets.find(s => s.properties.title === sheetTitle).properties.sheetId
      };
    }
    
    return structure;
  });
}

function prepareWriteOperations(collectedData, targetStructure) {
  const operations = [];
  
  for (const data of collectedData) {
    for (const [sheetTitle, structure] of Object.entries(targetStructure)) {
      const { isAllTestCases, firstColumn, secondColumn } = structure;
      
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
        operations.push({
          type: 'update',
          sheetTitle,
          rowIndex: existingC3Index,
          data,
          isAllTestCases
        });
      } else if (lastC24Index !== -1) {
        operations.push({
          type: 'insert',
          sheetTitle,
          afterRowIndex: lastC24Index,
          data,
          isAllTestCases
        });
      }
    }
  }
  
  return operations;
}

async function executeWriteOperations(auth, operations) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`📝 WRITE PHASE: ${operations.length} operations`);
  console.log(`${'='.repeat(70)}\n`);
  
  await cooldownWithProgress(RATE_LIMITS.BEFORE_WRITE_PHASE, 'Preparing for write phase');
  
  const sheets = google.sheets({ version: 'v4', auth });
  let successCount = 0;
  let failCount = 0;
  
  // Group operations by type for efficiency
  const insertOps = operations.filter(op => op.type === 'insert');
  const updateOps = operations.filter(op => op.type === 'update');
  
  // Process inserts first (they modify row indices)
  console.log(`\n📌 Processing ${insertOps.length} INSERT operations...`);
  for (let i = 0; i < insertOps.length; i++) {
    const op = insertOps[i];
    console.log(`   [${i + 1}/${insertOps.length}] Inserting in '${op.sheetTitle}' after row ${op.afterRowIndex}`);
    
    try {
      await withRetry(async () => {
        // Insert row
        await trackApiCall();
        const structure = await getSheetId(auth, op.sheetTitle);
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: {
            requests: [{
              insertDimension: {
                range: {
                  sheetId: structure,
                  dimension: 'ROWS',
                  startIndex: op.afterRowIndex,
                  endIndex: op.afterRowIndex + 1
                },
                inheritFromBefore: true
              }
            }]
          }
        });
        await sleep(RATE_LIMITS.BETWEEN_WRITES);
        
        // Write data
        await insertDataInRow(auth, sheets, op.sheetTitle, op.afterRowIndex + 1, op.data, op.isAllTestCases);
      });
      
      successCount++;
      console.log(`      ✅ Success`);
    } catch (error) {
      failCount++;
      console.error(`      ❌ Failed: ${error.message}`);
    }
    
    if (i < insertOps.length - 1) {
      await sleep(RATE_LIMITS.BETWEEN_WRITES);
    }
  }
  
  if (insertOps.length > 0 && updateOps.length > 0) {
    await cooldownWithProgress(RATE_LIMITS.BETWEEN_BATCH_WRITES, 'Between insert and update phases');
  }
  
  // Process updates
  console.log(`\n🔄 Processing ${updateOps.length} UPDATE operations...`);
  for (let i = 0; i < updateOps.length; i++) {
    const op = updateOps[i];
    console.log(`   [${i + 1}/${updateOps.length}] Updating '${op.sheetTitle}' row ${op.rowIndex}`);
    
    try {
      await withRetry(async () => {
        // Clear existing data
        await trackApiCall();
        const range = op.isAllTestCases ? `D${op.rowIndex}:T${op.rowIndex}` : `C${op.rowIndex}:S${op.rowIndex}`;
        await sheets.spreadsheets.values.clear({
          spreadsheetId: SHEET_ID,
          range: `${op.sheetTitle}!${range}`
        });
        await sleep(RATE_LIMITS.BETWEEN_WRITES);
        
        // Write new data
        await insertDataInRow(auth, sheets, op.sheetTitle, op.rowIndex, op.data, op.isAllTestCases);
      });
      
      successCount++;
      console.log(`      ✅ Success`);
    } catch (error) {
      failCount++;
      console.error(`      ❌ Failed: ${error.message}`);
    }
    
    if (i < updateOps.length - 1) {
      await sleep(RATE_LIMITS.BETWEEN_WRITES);
    }
  }
  
  console.log(`\n${'='.repeat(70)}`);
  console.log(`📊 WRITE SUMMARY: ✅ ${successCount} success, ❌ ${failCount} failed`);
  console.log(`${'='.repeat(70)}\n`);
  
  return { successCount, failCount };
}

async function getSheetId(auth, sheetTitle) {
  await trackApiCall();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  await sleep(RATE_LIMITS.BETWEEN_READS);
  return res.data.sheets.find(s => s.properties.title === sheetTitle).properties.sheetId;
}

function columnToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
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

async function insertDataInRow(auth, sheets, sheetTitle, row, data, isAllTestCases) {
  await trackApiCall();
  const startCol = isAllTestCases ? 'C' : 'B';

  const rowValues = [
    data.C24, data.C3, `=HYPERLINK("${data.sheetUrl}", "${data.C4}")`,
    data.B27, data.C5, data.C6, data.C7, '',
    data.C11, data.C32, data.C15, data.C13, data.C14,
    data.C18, data.C19, data.C20, data.C21
  ];

  if (isAllTestCases) rowValues.push(data.C21);

  const startIndex = columnToNumber(startCol);
  const endCol = numberToColumn(startIndex + rowValues.length - 1);
  const range = `${sheetTitle}!${startCol}${row}:${endCol}${row}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [rowValues] }
  });
  await sleep(RATE_LIMITS.BETWEEN_WRITES);
}

async function updateTestCasesInLibrary() {
  console.log('='.repeat(70));
  console.log('📋 Boards Test Cases Updater - OPTIMIZED VERSION');
  console.log('='.repeat(70));
  console.log(`⏰ Started: ${new Date().toISOString()}`);
  console.log('='.repeat(70));
  
  const startTime = Date.now();
  const authClient = await auth.getClient();
  
  try {
    // PHASE 1: Fetch all URLs
    console.log('\n🔍 PHASE 1: FETCHING ALL URLs');
    console.log('='.repeat(70));
    const urlsWithIndices = await fetchAllUrls(authClient);
    
    if (!urlsWithIndices.length) {
      console.log('❌ No URLs found!');
      return;
    }
    
    console.log(`✅ Found ${urlsWithIndices.length} URLs\n`);
    await sleep(RATE_LIMITS.BETWEEN_SPREADSHEETS);
    
    // Group by spreadsheet ID
    const spreadsheetMap = new Map();
    for (const { url, rowIndex } of urlsWithIndices) {
      const spreadsheetId = extractSpreadsheetId(url);
      if (!spreadsheetId) {
        console.log(`⚠️  Invalid URL in row ${rowIndex}`);
        continue;
      }
      
      if (!spreadsheetMap.has(spreadsheetId)) {
        spreadsheetMap.set(spreadsheetId, { spreadsheetId, rows: [rowIndex] });
      } else {
        spreadsheetMap.get(spreadsheetId).rows.push(rowIndex);
      }
    }
    
    const uniqueSpreadsheets = Array.from(spreadsheetMap.values());
    console.log(`📊 Processing ${uniqueSpreadsheets.length} unique spreadsheets\n`);
    
    // PHASE 2: Collect all data
    console.log('\n📥 PHASE 2: COLLECTING DATA FROM ALL SPREADSHEETS');
    console.log('='.repeat(70));
    const allCollectedData = [];
    let timeoutReached = false;
    
    for (let i = 0; i < uniqueSpreadsheets.length; i++) {
      const { spreadsheetId } = uniqueSpreadsheets[i];
      
      const result = await collectDataFromSpreadsheet(
        spreadsheetId, 
        authClient, 
        i + 1, 
        uniqueSpreadsheets.length, 
        startTime
      );
      
      if (result.timeout) {
        timeoutReached = true;
        console.log('⏰ Timeout reached during data collection');
        break;
      }
      
      if (result.data && result.data.length > 0) {
        allCollectedData.push(...result.data);
      }
      
      if (i < uniqueSpreadsheets.length - 1) {
        await sleep(RATE_LIMITS.BETWEEN_SPREADSHEETS);
      }
    }
    
    console.log(`\n✅ Data collection complete: ${allCollectedData.length} entries collected\n`);
    
    if (allCollectedData.length === 0) {
      console.log('❌ No data collected to write!');
      return;
    }
    
    // PHASE 3: Get target structure
    console.log('\n📋 PHASE 3: READING TARGET SHEET STRUCTURE');
    console.log('='.repeat(70));
    const targetStructure = await getTargetSheetStructure(authClient);
    console.log(`✅ Target structure loaded for ${Object.keys(targetStructure).length} sheets\n`);
    
    // PHASE 4: Prepare write operations
    console.log('\n🔨 PHASE 4: PREPARING WRITE OPERATIONS');
    console.log('='.repeat(70));
    const operations = prepareWriteOperations(allCollectedData, targetStructure);
    console.log(`✅ Prepared ${operations.length} write operations\n`);
    
    if (operations.length === 0) {
      console.log('❌ No operations to perform!');
      return;
    }
    
    // PHASE 5: Execute writes
    console.log('\n✍️  PHASE 5: EXECUTING WRITE OPERATIONS');
    const writeResults = await executeWriteOperations(authClient, operations);
    
    // Final summary
    const totalTime = Date.now() - startTime;
    console.log('\n' + '='.repeat(70));
    console.log('🎯 FINAL SUMMARY');
    console.log('='.repeat(70));
    console.log(`URLs processed: ${uniqueSpreadsheets.length}`);
    console.log(`Data entries collected: ${allCollectedData.length}`);
    console.log(`Write operations: ${operations.length}`);
    console.log(`✅ Successful writes: ${writeResults.successCount}`);
    console.log(`❌ Failed writes: ${writeResults.failCount}`);
    if (timeoutReached) console.log(`⏰ Timeout reached during collection`);
    console.log(`⏱️  Total runtime: ${formatTime(totalTime)}`);
    console.log('='.repeat(70));
    
  } catch (error) {
    console.error('\n❌ FATAL ERROR:', error);
    throw error;
  }
  
  console.log(`\n⏰ Finished: ${new Date().toISOString()}`);
}

updateTestCasesInLibrary().catch(console.error);
