function lossSummary() {
  // Spreadsheet ID for Profit Margin (PM) Spreadsheet.
  // Actual ID removed.
  let ss_id = '***';

  // Spreadsheet object for PM.
  let spreadsheet = SpreadsheetApp.openById(ss_id);

  // Array of all sheets in PM.
  let sheets = spreadsheet.getSheets();

  // Current spread sheet, Loss Summary.
  let lossSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Array for holding losses.
  let losses = [];

  // Access sheetCursor and rowCursor from Script Properties.
  let scriptProperties = PropertiesService.getScriptProperties();
  let sheetCursor = scriptProperties.getProperty('sheetCursor');
  let rowCursor = scriptProperties.getProperty('rowCursor');
 
  // If script properties are null, set them to the 5th sheet (6) and 2nd row (0-indexed).
  if (sheetCursor == null) {
    sheetCursor = 6;
    rowCursor = 1; 
  }

  // Parse cursors to ints.
  sheetCursor = parseInt(sheetCursor);
  rowCursor = parseInt(rowCursor);

  //Fetch data from the next possible value (rowCursor + 1) on sheetCursor.
  let currentSheet = sheets[sheetCursor].getDataRange().getValues();
  let nextRow = currentSheet[rowCursor + 1];

  // If nextRow on sheetCursor is empty, check the 2nd row next sheet.
  if (nextRow[0] == '') {
    let nextSheet = sheets[sheetCursor + 1].getDataRange().getValues();
    nextRow = nextSheet[1];

    //If the nextRow is not null, increment sheetCursor and set rowCursor to 0.
    if (nextRow[0] != '') {
      sheetCursor += 1;
      rowCursor = 0;
      currentSheet = nextSheet.slice();
    }
  }

  // While there is another value with a valid SO column, continue checking rows.
  while (nextRow[0] != '') {
    // Process the row.
    let newResult = processRow(nextRow);

    // If the row has a profit margin below 10%, add to losses array.
    if (newResult['margin'] <= 0.1) {
      losses.push(newResult);
    }
    // Update rowCursor, nextRow.
    rowCursor += 1;
    nextRow = currentSheet[rowCursor + 1];

    // If nextRow on sheetCursor is empty, check the 2nd row next sheet.
    if (nextRow[0] == '') {
      let nextSheet = sheets[sheetCursor + 1].getDataRange().getValues();
      nextRow = nextSheet[1];

      // If the nextRow is not null, increment sheetCursor and set rowCursor to 1.
      if (nextRow[0] != '') {
        sheetCursor += 1;
        rowCursor = 0;
        currentSheet = nextSheet.slice();
      }
    }
  }

  // Set the script properties to their new values.
  scriptProperties.setProperty('sheetCursor', sheetCursor.toString());
  scriptProperties.setProperty('rowCursor', rowCursor.toString());

  // Set starting row to write results based on values in sheet.
  let startRow = lossSheet.getDataRange().getNumRows() + 1;
  
  // Write results to spreadsheet.
  for (let i = 0; i < losses.length; i++) {
    let row = startRow + i;
    let obj = losses[i];

    // Populate SO, company, and notes columns.
    lossSheet.getRange(row, 1).setValue(obj['so']);
    lossSheet.getRange(row, 2).setValue(obj['company']);
    lossSheet.getRange(row, 8).setValue(obj['notes']);

    // Populate margin column as a percentage.
    lossSheet.getRange(row, 3).setValue(obj['margin']);
  }
}

// This function calculates the loss percent for each row.
// If the loss is negative, the row is added to the current spreadsheet.

function processRow(row) {
  // Result to return.
  let result = {};

  // Key values from each row.
  let company = row[1];
  let so = row[2];
  let billed = row[5];
  let notes = row[7];

  // Calculate total costs from values in the cost table.
  let cost = 0;
  let firstCostCol = 8;
  let lastCostCol = 17;
  for (let i = firstCostCol; i <= lastCostCol; i++) {
    if (row[i] != '') {
      cost += parseFloat(row[i]);
    }
  }
  
  // Calculate the margin.
  let margin = (billed - cost) / billed;

  // Construct return object.
  result['company'] = company;
  result['so'] = so;
  result['margin'] = margin;
  result['notes'] = notes;

  return result;
}
