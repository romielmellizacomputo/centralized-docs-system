while (row < rows.length) {
  const absRow = row + startRow;
  const fCell = (rows[row][1] || '').trim();
  const mergeEnd = mergedMap.get(absRow);
  const isMerged = mergeEnd && mergeEnd > absRow;
  const mergeLength = isMerged ? mergeEnd - absRow : 1;

  if (isMerged && !fCell) {
    // Add black borders for empty merged range (even though now unmerged)
    requests.push({
      updateBorders: {
        range: {
          sheetId: sheetMeta.properties.sheetId,
          startRowIndex: absRow,
          endRowIndex: mergeEnd,
          startColumnIndex: 4,
          endColumnIndex: 6
        },
        top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
        right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } }
      }
    });
  }

  if (fCell) {
    values[row] = [number.toString()];
    if (isMerged) {
      requests.push({
        mergeCells: {
          range: {
            sheetId: sheetMeta.properties.sheetId,
            startRowIndex: absRow,
            endRowIndex: mergeEnd,
            startColumnIndex: 4,
            endColumnIndex: 5,
          },
          mergeType: 'MERGE_ALL'
        }
      });
    }
    number++;
  }

  row += mergeLength;
}
