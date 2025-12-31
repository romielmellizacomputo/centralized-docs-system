import { google } from 'googleapis';
import { GoogleAuth } from 'google-auth-library';
import dotenv from 'dotenv';

dotenv.config();

const SHEET_ID = process.env.CDS_PORTAL_SPREADSHEET_ID;
const SHEET_NAME = 'Boards Test Cases';
const SHEETS_TO_SKIP = ['ToC', 'Roster', 'Issues', "HELP"];

// Progress tracking
const PROGRESS_TRACKER_CELL = 'J1';

// Batch processing configuration
const BATCH_SIZE = parseInt(process.env.BATCH_SIZE || '50');
const BATCH_NUMBER = parseInt(process.env.BATCH_NUMBER || '0');
const AUTO_INCREMENT = process.env.AUTO_INCREMENT === 'true';
const PROCESS_ALL_BATCHES = process.env.PROCESS_ALL_BATCHES === 'true';
const MAX_RUNTIME = 1.8 * 60 * 60 * 1000; // 1.8 hours

// Optimized rate limiting configuration
const RATE_LIMITS = {
  BETWEEN_SHEETS: 2000,
  BETWEEN_URLS: 5000,
  BETWEEN_OPERATIONS: 300,
  QUOTA_EXCEEDED_WAIT: 60000,
  BETWEEN_BATCHES: 10000
};

const auth = new GoogleAuth({
  credentials: JSON.parse(process.env.CDS_PORTALS_SERVICE_ACCOUNT_JSON),
  scopes: ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
});

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function formatTime(ms) {
  const seconds = Math.floor(ms / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);
  const remainingMinutes = minutes % 60;
  const remainingSeconds = seconds % 60;
  
  if (hours > 0) {
    return `${hours}h ${remainingMinutes}m ${remainingSeconds}s`;
  }
  if (minutes > 0) {
    return `${minutes}m ${remainingSeconds}s`;
  }
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

async function getLastProcessedBatch(auth) {
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!${PROGRESS_TRACKER_CELL}`
    });
    const value = response.data.values?.[0]?.[0];
    
    if (value && !isNaN(parseInt(value))) {
      const lastBatch = parseInt(value);
      console.log(`   📊 Last processed batch found: ${lastBatch}`);
      return lastBatch;
    }
    
    console.log('   📊 No previous batch tracked, starting from -1');
    return -1;
  } catch (error) {
    console.log('   ⚠️  Could not read progress tracker, starting from -1');
    return -1;
  }
}

async function setLastProcessedBatch(auth, batchNumber) {
  try {
    const sheets = google.sheets({ version: 'v4', auth });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!${PROGRESS_TRACKER_CELL}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[batchNumber]] }
    });
    console.log(`   💾 Progress saved: Batch ${batchNumber} completed`);
  } catch (error) {
    console.error('   ❌ Failed to save progress:', error.message);
  }
}

async function fetchUrls(auth, batchToProcess) {
  const sheets = google.sheets({ version: 'v4', auth });
  const range = `${SHEET_NAME}!D3:D`;
  
  console.log(`   🔍 Fetching URLs from ${range}...`);
  const response = await sheets.spreadsheets.values.get({ 
    spreadsheetId: SHEET_ID, 
    range 
  });
  const values = response.data.values || [];

  console.log(`   📥 Retrieved ${values.length} total rows from sheet`);
  
  const startIdx = batchToProcess * BATCH_SIZE;
  const endIdx = Math.min(startIdx + BATCH_SIZE, values.length);
  
  console.log(`   📦 Processing batch ${batchToProcess + 1}: rows ${startIdx + 3} to ${endIdx + 2} (${endIdx - startIdx} rows)`);
  
  if (startIdx >= values.length) {
    console.log(`   ⚠️  Batch ${batchToProcess} is beyond available data. No rows to process.`);
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
        console.log(`   ⏭️  No text found for cell D${actualRowIndex}`);
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
        urls.push({ url: hyperlink, rowIndex: actualRowIndex });
        console.log(`   ✅ Found URL in cell D${actualRowIndex}`);
        continue;
      }

      const urlRegex = /https?:\/\/\S+/;
      const match = text.match(urlRegex);
      const url = match ? match[0] : null;

      if (!url) {
        console.log(`   ⚠️  No URL found in text for cell D${actualRowIndex}`);
        continue;
      }

      urls.push({ url, rowIndex: actualRowIndex });
      console.log(`   ✅ Found URL in cell D${actualRowIndex}`);
    } catch (error) {
      console.error(`   ❌ Error processing URL for cell D${actualRowIndex}:`, error.message);
    }
  }

  return urls.filter(entry => entry && entry.url);
}

async function logData(auth, message) {
  const sheets = google.sheets({ version: 'v4', auth });
  const logCell = 'I1';
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

async function processUrl(spreadsheetId, auth, urlIndex, totalUrls, startTime) {
  const elapsed = Date.now() - startTime;
  if (elapsed > MAX_RUNTIME) {
    console.log('⏰ Approaching timeout limit, stopping gracefully...');
    return { timeout: true };
  }
  
  console.log(`\n${'='.repeat(70)}`);
  console.log(`🔄 Processing Spreadsheet ${urlIndex}/${totalUrls}`);
  console.log(`   📄 Spreadsheet ID: ${spreadsheetId}`);
  console.log(`   ⏱️  Elapsed time: ${formatTime(elapsed)}`);
  console.log(`${'='.repeat(70)}`);
  
  const sheets = google.sheets({ version: 'v4', auth });

  const meta = await sheets.spreadsheets.get({ spreadsheetId: spreadsheetId });
  const sheetTitles = meta.data.sheets
    .map(s => s.properties.title)
    .filter(title => !SHEETS_TO_SKIP.includes(title));

  console.log(`   📊 Found ${sheetTitles.length} sheets to process`);

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
  return { timeout: false };
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

async function processSingleBatch(authClient, startTime, forceBatchNumber = null) {
  let batchToProcess;
  
  if (forceBatchNumber !== null) {
    batchToProcess = forceBatchNumber;
  } else if (AUTO_INCREMENT) {
    const lastBatch = await getLastProcessedBatch(authClient);
    batchToProcess = lastBatch + 1;
    console.log(`🔄 Auto-increment mode enabled`);
    console.log(`   Last completed batch: ${lastBatch}`);
    console.log(`   Processing batch: ${batchToProcess}`);
  } else {
    batchToProcess = BATCH_NUMBER;
    console.log(`📌 Processing specified batch: ${batchToProcess}`);
  }
  
  console.log(`📦 Batch size: ${BATCH_SIZE} URLs`);
  console.log('='.repeat(70));
  console.log();
  
  console.log(`📥 Fetching URLs from sheet "${SHEET_NAME}"...`);
  const urlsWithIndices = await fetchUrls(authClient, batchToProcess);

  if (!urlsWithIndices.length) {
    await logData(authClient, `Batch ${batchToProcess}: No URLs to process.`);
    console.log('\n⚠️  No URLs found to process in this batch.');
    
    if (AUTO_INCREMENT && forceBatchNumber === null) {
      console.log('✅ All batches completed! Resetting to batch 0 for next cycle.');
      await setLastProcessedBatch(authClient, -1);
    }
    return { successCount: 0, failCount: 0, timeoutReached: false };
  }

  console.log(`\n✅ Found ${urlsWithIndices.length} URL(s) to process in this batch\n`);

  const spreadsheetMap = new Map();
  
  for (const { url, rowIndex } of urlsWithIndices) {
    const spreadsheetId = extractSpreadsheetId(url);
    
    if (!spreadsheetId) {
      console.log(`⚠️  Could not extract spreadsheet ID from URL in row D${rowIndex}`);
      await logData(authClient, `Batch ${batchToProcess}: Invalid URL format in row D${rowIndex}: ${url}`);
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
  console.log(`   Total URLs found in batch: ${totalUrls}`);
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

  await logData(authClient, `Batch ${batchToProcess}: Found ${totalUrls} URL(s), ${uniqueSpreadsheets.length} unique spreadsheet(s) to process`);

  let successCount = 0;
  let failCount = 0;
  let timeoutReached = false;

  for (let i = 0; i < uniqueSpreadsheets.length; i++) {
    const { spreadsheetId, url, rows } = uniqueSpreadsheets[i];
    
    const rowsText = rows.length > 1 
      ? `rows D${rows.join(', D')}` 
      : `row D${rows[0]}`;
    
    await logData(authClient, `Batch ${batchToProcess}: Processing spreadsheet ${spreadsheetId} (from ${rowsText})`);
    
    try {
      const result = await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length, startTime);
      
      if (result.timeout) {
        timeoutReached = true;
        await logData(authClient, `Batch ${batchToProcess}: Stopped at ${i + 1}/${uniqueSpreadsheets.length} due to approaching timeout`);
        break;
      }
      
      successCount++;
    } catch (error) {
      failCount++;
      
      if (error.message.includes('Quota exceeded') || error.code === 429) {
        await logData(authClient, `⚠️  Quota exceeded for spreadsheet: ${spreadsheetId}. Retrying after cooldown...`);
        await cooldownWithProgress(RATE_LIMITS.QUOTA_EXCEEDED_WAIT, 'Quota exceeded - cooling down');
        
        try {
          const result = await processUrl(spreadsheetId, authClient, i + 1, uniqueSpreadsheets.length, startTime);
          
          if (result.timeout) {
            timeoutReached = true;
            await logData(authClient, `Batch ${batchToProcess}: Stopped at ${i + 1}/${uniqueSpreadsheets.length} due to approaching timeout`);
            break;
          }
          
          successCount++;
          failCount--;
        } catch (retryError) {
          await logData(authClient, `❌ Error processing spreadsheet on retry: ${spreadsheetId}. Error: ${retryError.message}`);
        }
      } else {
        await logData(authClient, `❌ Error processing spreadsheet: ${spreadsheetId}. Error: ${error.message}`);
      }
    }

    if (i < uniqueSpreadsheets.length - 1 && !timeoutReached) {
      await cooldownWithProgress(
        RATE_LIMITS.BETWEEN_URLS, 
        `Cooling down before next spreadsheet (${i + 1}/${uniqueSpreadsheets.length} complete)`
      );
    }
  }

  if (AUTO_INCREMENT && successCount > 0 && !timeoutReached && forceBatchNumber === null) {
    await setLastProcessedBatch(authClient, batchToProcess);
  }

  const totalTime = Date.now() - startTime;
  
  console.log('\n' + '='.repeat(70));
  console.log('  📊 BATCH SUMMARY');
  console.log('='.repeat(70));
  console.log(`📦 Batch number: ${batchToProcess + 1}`);
  console.log(`📄 Total URLs found in batch: ${totalUrls}`);
  console.log(`📚 Unique spreadsheets in batch: ${uniqueSpreadsheets.length}`);
  if (duplicateCount > 0) {
    console.log(`🔗 Duplicate references detected: ${duplicateCount}`);
  }
  console.log(`✅ Successful: ${successCount}/${uniqueSpreadsheets.length}`);
  if (failCount > 0) {
    console.log(`❌ Failed: ${failCount}/${uniqueSpreadsheets.length}`);
  }
  if (timeoutReached) {
    console.log(`⏰ Stopped early due to timeout safety limit`);
  }
  console.log(`⏱️  Batch time: ${formatTime(totalTime)}`);
  console.log('='.repeat(70));

  await logData(authClient, `Batch ${batchToProcess}: Complete. Success: ${successCount}, Failed: ${failCount}`);
  
  return { successCount, failCount, timeoutReached };
}

async function processAllBatches(authClient, startTime) {
  console.log('🔄 PROCESSING ALL BATCHES MODE');
  console.log('📊 Fetching total number of URLs...');
  
  const sheets = google.sheets({ version: 'v4', auth: authClient });
  const range = `${SHEET_NAME}!D3:D`;
  const response = await sheets.spreadsheets.values.get({ 
    spreadsheetId: SHEET_ID, 
    range 
  });
  const totalUrls = (response.data.values || []).length;
  const totalBatches = Math.ceil(totalUrls / BATCH_SIZE);
  
  console.log(`📦 Total URLs: ${totalUrls}`);
  console.log(`📦 Batch size: ${BATCH_SIZE}`);
  console.log(`📦 Total batches to process: ${totalBatches}`);
  console.log('='.repeat(70));
  
  let totalSuccess = 0;
  let totalFail = 0;
  let batchesProcessed = 0;
  
  for (let batchNum = 0; batchNum < totalBatches; batchNum++) {
    const elapsed = Date.now() - startTime;
    if (elapsed > MAX_RUNTIME) {
      console.log(`⏰ Reached time limit after ${batchNum} batches`);
      await logData(authClient, `All-batch run: Stopped at batch ${batchNum}/${totalBatches} due to timeout`);
      break;
    }
    
    console.log(`\n${'*'.repeat(70)}`);
    console.log(`📦 PROCESSING BATCH ${batchNum + 1}/${totalBatches}`);
    console.log(`⏱️  Elapsed: ${formatTime(elapsed)} | Remaining: ${formatTime(MAX_RUNTIME - elapsed)}`);
    console.log(`${'*'.repeat(70)}\n`);
    
    const result = await processSingleBatch(authClient, startTime, batchNum);
    totalSuccess += result.successCount;
    totalFail += result.failCount;
    batchesProcessed++;
    
    if (result.timeoutReached) {
      console.log('⏰ Stopping all-batch processing due to timeout in current batch');
      break;
    }
    
    // Cooldown between batches
    if (batchNum < totalBatches - 1) {
      await cooldownWithProgress(RATE_LIMITS.BETWEEN_BATCHES, `Cooling down between batches`);
    }
  }
  
  // Save final progress
  if (AUTO_INCREMENT && batchesProcessed > 0) {
    await setLastProcessedBatch(authClient, batchesProcessed - 1);
  }
  
  const totalTime = Date.now() - startTime;
  console.log('\n' + '='.repeat(70));
  console.log('  🎯 FINAL SUMMARY - ALL BATCHES');
  console.log('='.repeat(70));
  console.log(`📦 Batches processed: ${batchesProcessed}/${totalBatches}`);
  console.log(`✅ Total successful: ${totalSuccess}`);
  console.log(`❌ Total failed: ${totalFail}`);
  console.log(`⏱️  Total time: ${formatTime(totalTime)}`);
  console.log('='.repeat(70));
  
  await logData(authClient, `All-batch run complete: ${batchesProcessed}/${totalBatches} batches, Success: ${totalSuccess}, Failed: ${totalFail}`);
}

async function updateTestCasesInLibrary() {
  console.log('='.repeat(70));
  console.log('  📋 Boards Test Cases Updater - Batch Processing');
  console.log('='.repeat(70));
  console.log(`⏰ Started at: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}`);
  
  const startTime = Date.now();
  const authClient = await auth.getClient();
  
  if (PROCESS_ALL_BATCHES) {
    console.log('🔄 Mode: PROCESS ALL BATCHES');
    await processAllBatches(authClient, startTime);
  } else {
    console.log('🔄 Mode: SINGLE BATCH');
    await processSingleBatch(authClient, startTime);
  }
  
  const totalTime = Date.now() - startTime;
  console.log(`\n⏰ Finished at: ${new Date().toISOString().replace('T', ' ').substring(0, 19)}`);
  console.log(`⏱️  Total runtime: ${formatTime(totalTime)}`);
}

updateTestCasesInLibrary().catch(console.error);
