import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues', "HELP"];
const PROGRESS_TRACKER_CELL = 'J1';
const BATCH_SIZE = parseInt(process.env.BATCH_SIZE || '30'); // REDUCED from 50
const BATCH_NUMBER = parseInt(process.env.BATCH_NUMBER || '0');
const AUTO_INCREMENT = process.env.AUTO_INCREMENT === 'true';
const PROCESS_ALL_BATCHES = process.env.PROCESS_ALL_BATCHES === 'true';
const MAX_RUNTIME = 1.8 * 60 * 60 * 1000;

// AGGRESSIVE rate limiting
const RATE_LIMITS = {
  BETWEEN_SHEETS: 4000,
  BETWEEN_URLS: 10000,
  BETWEEN_OPERATIONS: 600,
  QUOTA_EXCEEDED_WAIT: 120000,
  BETWEEN_BATCHES: 20000,
  BETWEEN_VALIDATIONS: 3000,
  AFTER_BATCH_GET: 1000,
  AFTER_WRITE: 1500
};

// Track API calls per minute
let apiCallCount = 0;
let lastResetTime = Date.now();
const MAX_CALLS_PER_MINUTE = 50; // Conservative limit
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
    console.log(`   📊 Resetting API counter. Previous minute: ${apiCallCount} calls`);
    apiCallCount = 1;
    lastResetTime = now;
    return;
  }
  
  if (apiCallCount >= MAX_CALLS_PER_MINUTE - 3) {
    const waitTime = MINUTE - (now - lastResetTime) + 2000; // Extra 2s buffer
    console.log(`   ⚠️  Rate limit protection: ${apiCallCount} calls made, pausing for ${formatTime(waitTime)}`);
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
  
  console.log(`\r   ✅ Cooldown complete!${' '.repeat(50)}`);
}

async function withRetry(fn, maxRetries = 2, baseDelay = RATE_LIMITS.QUOTA_EXCEEDED_WAIT) {
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

async function getLastProcessedBatch(auth) {
  try {
    return await withRetry(async () => {
      await trackApiCall();
      const sheets = google.sheets({ version: 'v4', auth });
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${SHEET_NAME}!${PROGRESS_TRACKER_CELL}`
      });
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      const value = response.data.values?.[0]?.[0];
      if (value && !isNaN(parseInt(value))) {
        console.log(`   📊 Last processed batch: ${parseInt(value)}`);
        return parseInt(value);
      }
      console.log('   📊 No previous batch tracked, starting from -1');
      return -1;
    });
  } catch (error) {
    console.log('   ⚠️  Could not read progress tracker, starting from -1');
    return -1;
  }
}

async function setLastProcessedBatch(auth, batchNumber) {
  try {
    await withRetry(async () => {
      await trackApiCall();
      const sheets = google.sheets({ version: 'v4', auth });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${SHEET_NAME}!${PROGRESS_TRACKER_CELL}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[batchNumber]] }
      });
      await sleep(RATE_LIMITS.AFTER_WRITE);
      console.log(`   💾 Progress saved: Batch ${batchNumber}`);
    });
  } catch (error) {
    console.error('   ❌ Failed to save progress:', error.message);
  }
}

async function fetchUrls(auth, batchToProcess) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const range = `${SHEET_NAME}!D3:D`;
    
    console.log(`   🔍 Fetching URLs from ${range}...`);
    const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    
    const values = response.data.values || [];
    console.log(`   📥 Retrieved ${values.length} total rows`);
    
    const startIdx = batchToProcess * BATCH_SIZE;
    const endIdx = Math.min(startIdx + BATCH_SIZE, values.length);
    
    console.log(`   📦 Batch ${batchToProcess + 1}: rows ${startIdx + 3} to ${endIdx + 2} (${endIdx - startIdx} rows)`);
    
    if (startIdx >= values.length) {
      console.log(`   ⚠️  Batch ${batchToProcess} is beyond available data`);
      return [];
    }
    
    const batchValues = values.slice(startIdx, endIdx);
    const urls = [];
    
    for (let index = 0; index < batchValues.length; index++) {
      const row = batchValues[index];
      const actualRowIndex = startIdx + index + 3;
      const cellRange = `${SHEET_NAME}!D${actualRowIndex}`;

      try {
        const text = row[0] || null;
        if (!text) {
          console.log(`   ⏭️  No text in D${actualRowIndex}`);
          continue;
        }

        await trackApiCall();
        const linkResponse = await sheets.spreadsheets.get({
          spreadsheetId: SHEET_ID,
          ranges: [cellRange],
          includeGridData: true,
          fields: 'sheets.data.rowData.values.hyperlink'
        });

        await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);

        const hyperlink = linkResponse.data.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0]?.hyperlink;

        if (hyperlink) {
          urls.push({ url: hyperlink, rowIndex: actualRowIndex });
          console.log(`   ✅ Found URL in D${actualRowIndex}`);
          continue;
        }

        const urlRegex = /https?:\/\/\S+/;
        const match = text.match(urlRegex);
        const url = match ? match[0] : null;

        if (!url) {
          console.log(`   ⚠️  No URL in D${actualRowIndex}`);
          continue;
        }

        urls.push({ url, rowIndex: actualRowIndex });
        console.log(`   ✅ Found URL in D${actualRowIndex}`);
      } catch (error) {
        console.error(`   ❌ Error at D${actualRowIndex}: ${error.message}`);
      }
    }

    return urls.filter(entry => entry && entry.url);
  });
}

async function logData(auth, message) {
  try {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const timestamp = new Date().toISOString().replace('T', ' ').substring(0, 19);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!I1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[`[${timestamp}] ${message}`]] }
    });
    console.log(`   📋 ${message}`);
    await sleep(RATE_LIMITS.AFTER_WRITE);
  } catch (error) {
    console.error(`   ⚠️  Logging failed: ${error.message}`);
  }
}

async function collectSheetData(auth, spreadsheetId, sheetTitle) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const cellRefs = ['C3', 'C4', 'C5', 'C6', 'C7', 'C13', 'C14', 'C15', 'C18', 'C19', 'C20', 'C21', 'C24', 'B27', 'C32', 'C11'];
    const ranges = cellRefs.map(ref => `${sheetTitle}!${ref}`);

    const res = await sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges });
    await sleep(RATE_LIMITS.AFTER_BATCH_GET);

    const data = {};
    res.data.valueRanges.forEach((range, index) => {
      data[cellRefs[index]] = range.values?.[0]?.[0] || null;
    });

    if (!data['C24']) return null;

    await trackApiCall();
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    
    const sheet = meta.data.sheets.find(s => s.properties.title === sheetTitle);
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.properties.sheetId}`;

    return { ...data, sheetUrl, sheetName: sheetTitle };
  });
}

async function processUrl(spreadsheetId, auth, urlIndex, totalUrls, startTime) {
  const elapsed = Date.now() - startTime;
  if (elapsed > MAX_RUNTIME) {
    console.log('⏰ Timeout limit approaching');
    return { timeout: true };
  }
  
  console.log(`\n${'='.repeat(70)}`);
  console.log(`🔄 Spreadsheet ${urlIndex}/${totalUrls} - ID: ${spreadsheetId}`);
  console.log(`   ⏱️  Elapsed: ${formatTime(elapsed)} | API calls: ${apiCallCount}`);
  console.log(`${'='.repeat(70)}`);
  
  try {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    
    const sheetTitles = meta.data.sheets
      .map(s => s.properties.title)
      .filter(title => !SHEETS_TO_SKIP.includes(title));

    console.log(`   📊 Found ${sheetTitles.length} sheets`);

    const validData = [];
    for (let i = 0; i < sheetTitles.length; i++) {
      console.log(`   🔄 [${i + 1}/${sheetTitles.length}] ${sheetTitles[i]}`);
      
      const data = await collectSheetData(auth, spreadsheetId, sheetTitles[i]);
      
      if (data && Object.values(data).some(v => v !== null && v !== '')) {
        validData.push(data);
        console.log(`      ✅ Valid data`);
      } else {
        console.log(`      ⏭️  No data`);
      }
      
      if (i < sheetTitles.length - 1) {
        await cooldownWithProgress(RATE_LIMITS.BETWEEN_SHEETS, 'Before next sheet');
      }
    }

    console.log(`\n   📦 Processing ${validData.length} entries...`);

    for (let i = 0; i < validData.length; i++) {
      console.log(`   🔄 [${i + 1}/${validData.length}] ${validData[i].sheetName}`);
      await logData(auth, `Processing: ${validData[i].sheetName}`);
      await validateAndInsertData(auth, validData[i]);
      
      if (i < validData.length - 1) {
        await sleep(RATE_LIMITS.BETWEEN_VALIDATIONS);
      }
    }

    console.log(`   ✅ Complete`);
    return { timeout: false };
  } catch (error) {
    console.error(`   ❌ Error: ${error.message}`);
    throw error;
  }
}

async function validateAndInsertData(auth, data) {
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
      await sleep(RATE_LIMITS.AFTER_WRITE);
      
      await logData(auth, `Updated row ${existingC3Index} in '${sheetTitle}'`);
      processed = true;
    } else if (lastC24Index !== -1) {
      const newRowIndex = lastC24Index + 1;
      await insertRowWithFormat(auth, sheetTitle, lastC24Index);
      await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
      
      await insertDataInRow(auth, sheetTitle, newRowIndex, data, isAllTestCases ? 'C' : 'B');
      await sleep(RATE_LIMITS.AFTER_WRITE);
      
      await logData(auth, `Inserted row after ${lastC24Index} in '${sheetTitle}'`);
      processed = true;
    }

    await sleep(RATE_LIMITS.BETWEEN_SHEETS);
  }

  if (!processed) {
    await logData(auth, `No matches for C24='${data.C24}' or C3='${data.C3}'`);
  }

  return processed;
}

async function insertRowWithFormat(auth, sheetTitle, sourceRowIndex) {
  await withRetry(async () => {
    await trackApiCall();
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
    await sleep(RATE_LIMITS.AFTER_WRITE);
    console.log(`      📝 Inserted row after ${sourceRowIndex}`);
  });
}

async function getSheetId(auth, sheetTitle) {
  await trackApiCall();
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
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

async function insertDataInRow(auth, sheetTitle, row, data, startCol) {
  await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const isAllTestCases = sheetTitle === "ALL TEST CASES";

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
    await sleep(RATE_LIMITS.AFTER_WRITE);
  });
}

async function clearRowData(auth, sheetTitle, row, isAllTestCases) {
  await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const range = isAllTestCases ? `D${row}:T${row}` : `C${row}:S${row}`;
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SHEET_ID,
      range: `${sheetTitle}!${range}`
    });
    await sleep(RATE_LIMITS.AFTER_WRITE);
  });
}

async function getTargetSheetTitles(auth) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    return meta.data.sheets.map(s => s.properties.title);
  });
}

async function getColumnValues(auth, sheetTitle, column) {
  return await withRetry(async () => {
    await trackApiCall();
    const sheets = google.sheets({ version: 'v4', auth });
    const range = `${sheetTitle}!${column}:${column}`;
    const res = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range });
    await sleep(RATE_LIMITS.BETWEEN_OPERATIONS);
    return res.data.values?.map(row => row[0]) || [];
  });
}

async function processSingleBatch(authClient, startTime, forceBatchNumber = null) {
  let batchToProcess = forceBatchNumber !== null ? forceBatchNumber : 
    (AUTO_INCREMENT ? await getLastProcessedBatch(authClient) + 1 : BATCH_NUMBER);
  
  console.log(`📦 Processing batch ${batchToProcess}, size: ${BATCH_SIZE}`);
  console.log('='.repeat(70));
  
  const urlsWithIndices = await fetchUrls(authClient, batchToProcess);

  if (!urlsWithIndices.length) {
    await logData(authClient, `Batch ${batchToProcess}: No URLs`);
    if (AUTO_INCREMENT && forceBatchNumber === null) {
      console.log('✅ All batches done! Resetting.');
      await setLastProcessedBatch(authClient, -1);
    }
    return { successCount: 0, failCount: 0, timeoutReached: false };
  }

  console.log(`\n✅ Found ${urlsWithIndices.length} URL(s)\n`);

  const spreadsheetMap = new Map();
  for (const { url, rowIndex } of urlsWithIndices) {
    const spreadsheetId = extractSpreadsheetId(url);
    if (!spreadsheetId) {
      console.log(`⚠️  Invalid URL in D${rowIndex}`);
      continue;
    }
    
    if (!spreadsheetMap.has(spreadsheetId)) {
      spreadsheetMap.set(spreadsheetId, { url, rows: [rowIndex], spreadsheetId });
    } else {
      spreadsheetMap.get(spreadsheetId).rows.push(rowIndex);
    }
  }

  const uniqueSpreadsheets = Array.from(spreadsheetMap.values());
  console.log(`📊 ${urlsWithIndices.length} URLs, ${uniqueSpreadsheets.length} unique spreadsheets\n`);

  let successCount = 0, failCount = 0, timeoutReached = false;

  for (let i = 0; i < uniqueSpreadsheets.length; i++) {
    const { spreadsheetId, rows } = uniqueSpreadsheets[i];
    
    await logData(authClient, `Batch ${batchToProcess}: Processing ${spreadsheetId}`);
    
    try {
      const result = await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length, startTime);
      
      if (result.timeout) {
        timeoutReached = true;
        break;
      }
      successCount++;
    } catch (error) {
      failCount++;
      if (error.message.includes('Quota exceeded') || error.code === 429) {
        await logData(authClient, `⚠️  Quota exceeded: ${spreadsheetId}`);
        await cooldownWithProgress(RATE_LIMITS.QUOTA_EXCEEDED_WAIT, 'Quota recovery');
        
        try {
          const result = await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length, startTime);
          if (result.timeout) {
            timeoutReached = true;
            break;
          }
          successCount++;
          failCount--;
        } catch (retryError) {
          await logData(authClient, `❌ Retry failed: ${retryError.message}`);
        }
      } else {
        await logData(authClient, `❌ Error: ${error.message}`);
      }
    }

    if (i < uniqueSpreadsheets.length - 1 && !timeoutReached) {
      await cooldownWithProgress(RATE_LIMITS.BETWEEN_URLS, `Progress: ${i + 1}/${uniqueSpreadsheets.length}`);
    }
  }

  if (AUTO_INCREMENT && successCount > 0 && !timeoutReached && forceBatchNumber === null) {
    await setLastProcessedBatch(authClient, batchToProcess);
  }

  const totalTime = Date.now() - startTime;
  console.log('\n' + '='.repeat(70));
  console.log('📊 BATCH SUMMARY');
  console.log('='.repeat(70));
  console.log(`Batch: ${batchToProcess + 1}`);
  console.log(`✅ Success: ${successCount}/${uniqueSpreadsheets.length}`);
  if (failCount > 0) console.log(`❌ Failed: ${failCount}`);
  if (timeoutReached) console.log(`⏰ Stopped early (timeout)`);
  console.log(`⏱️  Time: ${formatTime(totalTime)}`);
  console.log('='.repeat(70));

  return { successCount, failCount, timeoutReached };
}

async function processAllBatches(authClient, startTime) {
  console.log('🔄 PROCESSING ALL BATCHES');
  
  await trackApiCall();
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const response = await sheets.spreadsheets.values.get({ 
    spreadsheetId: SHEET_ID, 
    range: `${SHEET_NAME}!D3:D`
  });
  const totalUrls = (response.data.values || []).length;
  const totalBatches = Math.ceil(totalUrls / BATCH_SIZE);
  
  console.log(`📦 Total: ${totalUrls} URLs, ${totalBatches} batches`);
  console.log('='.repeat(70));
  
  let totalSuccess = 0, totalFail = 0, batchesProcessed = 0;
  
  for (let batchNum = 0; batchNum < totalBatches; batchNum++) {
    const elapsed = Date.now() - startTime;
    if (elapsed > MAX_RUNTIME) {
      console.log(`⏰ Time limit at batch ${batchNum}`);
      break;
    }
    
    console.log(`\n${'*'.repeat(70)}`);
    console.log(`📦 BATCH ${batchNum + 1}/${totalBatches}`);
    console.log(`${'*'.repeat(70)}\n`);
    
    const result = await processSingleBatch(authClient, startTime, batchNum);
    totalSuccess += result.successCount;
    totalFail += result.failCount;
    batchesProcessed++;
    
    if (result.timeoutReached) break;
    
    if (batchNum < totalBatches - 1) {
      await cooldownWithProgress(RATE_LIMITS.BETWEEN_BATCHES, 'Between batches');
    }
  }
  
  if (AUTO_INCREMENT && batchesProcessed > 0) {
    await setLastProcessedBatch(authClient, batchesProcessed - 1);
  }
  
  console.log('\n' + '='.repeat(70));
  console.log('🎯 FINAL SUMMARY');
  console.log('='.repeat(70));
  console.log(`Batches: ${batchesProcessed}/${totalBatches}`);
  console.log(`✅ Success: ${totalSuccess}`);
  console.log(`❌ Failed: ${totalFail}`);
  console.log(`⏱️  Time: ${formatTime(Date.now() - startTime)}`);
  console.log('='.repeat(70));
}

async function updateTestCasesInLibrary() {
  console.log('='.repeat(70));
  console.log('📋 Boards Test Cases Updater');
  console.log('='.repeat(70));
  console.log(`⏰ Started: ${new Date().toISOString()}`);
  
  const startTime = Date.now();
  const authClient = await auth.getClient();
  
  if (PROCESS_ALL_BATCHES) {
    await processAllBatches(authClient, startTime);
  } else {
    await processSingleBatch(authClient, startTime);
  }
  
  console.log(`\n⏰ Finished: ${new Date().toISOString()}`);
  console.log(`⏱️  Runtime: ${formatTime(Date.now() - startTime)}`);
}

updateTestCasesInLibrary().catch(console.error);
