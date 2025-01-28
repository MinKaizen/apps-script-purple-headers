const packageName = 'Purple Headers'

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu(packageName)
    .addItem('Append', 'copyAndAppend')
    .addItem('Replace All', 'copyAndReplaceAll')
    .addToUi()
}

function isBlank(value) {
  return value == "" || value == null
}

/**
 * Checks whether all cells in a range have formulas
 * 
 * @param {Range} range
 * @return boolean
 */
function rangeAllCellsHaveFormulas(range) {
  const formulas = range.getFormulas();
  return !formulas.some(row => row.some(cell => isBlank(cell)));
}

/**
 * Given a range, get the absolute row number of the last row with
 * at least one non-blank cell
 * 
 * @param {Range} range 
 * @returns {number}
 */
function rangeGetLastRow(range) {
  const values = range.getValues();
  const isNonBlankRow = (row) => row.some(cell => !isBlank(cell));

  // Loop through the rows in reverse order
  for (let r = values.length - 1; r >= 0; r--) {
    // If the row has at least one non-blank cell, return the row index
    if (isNonBlankRow(values[r])) {
      // r is 0-based, but Rows are 1-based
      return range.getRow() + r;
    }
  }

  // If all rows are blank, return the first row -1
  return range.getRow() - 1; 
}

/**
 * Given a Range, return an array of Ranges where the values satisfy the callback
 * 
 * @param {Range} range
 * @param {function(any): boolean} callback
 * @returns {Range[]}
 */
function rangeMapValues(range, callback) {
  const sheet = range.getSheet();
  const values = range.getValues();
  const ranges = [];

  values.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (callback(value)) {
        const cellRange = sheet.getRange(range.getRow() + rowIndex, range.getColumn() + colIndex);
        ranges.push(cellRange);
      }
    });
  });

  return ranges;
}

/**
 * Given a Range, return an array of Ranges where the formulas satisfy the callback
 * 
 * @param {Range} range
 * @param {function(string): boolean} callback
 * @returns {Range[]}
 */
function rangeMapFormulas(range, callback) {
  const sheet = range.getSheet();
  const formulas = range.getFormulas();
  const ranges = [];

  formulas.forEach((row, rowIndex) => {
    row.forEach((formula, colIndex) => {
      if (callback(formula)) {
        const cellRange = sheet.getRange(range.getRow() + rowIndex, range.getColumn() + colIndex);
        ranges.push(cellRange);
      }
    });
  });

  return ranges;
}

/**
 * Clears a range of formulas, font style and font folor
 * 
 * @param {*} range 
 * @returns void
 */
function rangeClear(range) {
  const values = range.getValues()
  range.setValues(values)
  range.setFontColor(null)
  range.setFontStyle(null)
}

/**
 * Fills the rest of the rows with the selected formulas, skipping
 * gaps and rows with values (even if only one cell in a row has a value)
 * 
 * @returns void
 */
function copyAndAppend() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const selection = sheet.getActiveRange()
  const selectionRow = selection.getRow()
  const selectionColumn = selection.getColumn()
  const rowEnd = sheet.getLastRow()
  const rowStart = rangeGetLastRow(sheet.getRange(selectionRow, selectionColumn, rowEnd)) + 1
  const numRows = rowEnd - rowStart + 1
  applyFormulasDownard_(selection, rowStart, numRows)
}

/**
 * Copies formulas from a single row of cells down until the last row
 */
function copyAndReplaceAll() {
  const sheet = SpreadsheetApp.getActiveSheet()
  const selection = sheet.getActiveRange()
  const rowEnd = sheet.getLastRow()
  const rowStart = selection.getRow() + 1
  const numRows = rowEnd - rowStart + 1
  applyFormulasDownard_(selection, rowStart, numRows)
}

/**
 * Applies formulas from a row of cells downward, then clears the formulas, font style and font color
 * 
 * @param {*} range The range that contains the formulas to copy
 * @param {*} rowStart The row number to start copying the formulas to
 * @param {*} numRows The number of rows to copy the formulas to
 */
function applyFormulasDownard_(range, rowStart, numRows) {
  const sheet = range.getSheet()

  if (range.getNumRows() > 1) {
    throw new Error(`${packageName} doesn't work on multiple rows!`)
  }

  const rangeColumn = range.getColumn()
  const rangesToCopy = rangeMapFormulas(range, formula => formula !== "")

  for (let i = 0; i < rangesToCopy.length; i++) {
    const rangeToCopy = rangesToCopy[i]
    const column = rangeToCopy.getColumn()
    const rangeToPasteFormulas = sheet.getRange(rowStart, column, numRows)
    rangeToCopy.copyTo(rangeToPasteFormulas)
  }

  const rangeToClear = sheet.getRange(rowStart, rangeColumn, numRows, range.getNumColumns())
  rangeClear(rangeToClear)
}