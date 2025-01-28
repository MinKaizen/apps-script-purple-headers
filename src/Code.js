const packageName = 'Purple Headers'

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu(packageName)
    .addItem('Replace All', 'copyAndReplaceAll')
    .addToUi()
}

function isBlank(value) {
  return value == "" || value == null
}

/**
 * Checks all cells in a range have formulas
 * 
 * @param {Range} range
 * @return boolean
 */
function rangeAllCellsHaveFormulas(range) {
  const formulas = range.getFormulas();
  return !formulas.some(row => row.some(cell => isBlank(cell)));
}

/**
 * Copies formulas from a single row of cells down until the last row
 */
function copyAndReplaceAll() {
  const currentRange = SpreadsheetApp.getActiveRange()

  if (currentRange.getNumRows() > 1) {
    throw new Error(`${packageName} doesn't work on multiple rows!`)
  }

  const currentSheet = SpreadsheetApp.getActiveSheet()
  const sheetMaxRows = currentSheet.getLastRow()
  // const formulas = currentRange.getFormulas()[0]

  if (!rangeAllCellsHaveFormulas(currentRange)) {
    throw new Error('Not all selected cells have formulas!')
  }

  const targetRange = currentSheet.getRange(currentRange.getRow(), currentRange.getColumn(), sheetMaxRows, currentRange.getNumColumns())
  currentRange.copyTo(targetRange)
  const targetValuesRange = targetRange.offset(1, 0)
  const values = targetValuesRange.getValues()
  targetValuesRange.setValues(values)
  targetValuesRange.setFontColor(null)
  targetValuesRange.setFontStyle(null)
}
