const { GoogleSpreadsheet } = require('google-spreadsheet');

async function autoNumberSteps(sheetData) {
  const doc = new GoogleSpreadsheet(sheetData.sheetId);
  await doc.useServiceAccountAuth(JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON));
  
  await doc.loadInfo();
  const sheet = doc.sheetsByTitle[sheetData.sheetName];

  const rangeF = sheet.getRange("F12:F");
  const valuesF = await rangeF.getValues();
  const numRows = valuesF.length;

  let lastStepNumber = 0;
  let currentRow = 12;

  for (let i = 0; i < numRows; i++) {
    const cellF = sheet.getCell(currentRow - 1, 5); // Column F
    const cellE = sheet.getCell(currentRow - 1, 4); // Column E

    if (cellF.value) {
      if (!cellE.value) {
        lastStepNumber++;
        cellE.value = lastStepNumber;
      } else {
        const currentStep = cellE.value;
        if (currentStep !== lastStepNumber + 1) {
          cellE.value = lastStepNumber + 1;
        }
      }
      lastStepNumber = cellE.value;
      currentRow++;
    } else {
      if (cellE.value) {
        cellE.value = ''; // Clear E cell
      }
      currentRow++;
    }
  }

  await sheet.saveUpdatedCells();
}

const sheetData = JSON.parse(process.env.INPUT_SHEET_DATA);
autoNumberSteps(sheetData).catch(console.error);
