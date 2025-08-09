var spreadsheet = SpreadsheetApp.getActive();
var sheet = spreadsheet.getSheetByName("COTIZADOR")

function getEndRow(startRow) {
  if (startRow < 1) throw new Error("startRow must be >= 1");
  let row = startRow + 1;
  while (!isRowEmpty(row)) row++;
  return row - 1;
}

function isRowEmpty(rowNumber) {
 
  var rowValues = sheet.getRange(rowNumber, 1, 1, c_LAST_COLUMN_WITH_DATA).getValues()[0];
  // Check if every cell is empty
  return rowValues.every(function(cell) {
    return cell === "" || cell === null;
  });
}

function collapseRange(range, startRow) {
  range.shiftRowGroupDepth(1);
  group = sheet.getRowGroup(startRow, 1)
  group.collapse()
}

function getRangeSrc(){
  let lastRow = 0 

  if ( ReadCount() != 0 )
     lastRow = GetBaseLastRow()
  else{
    lastRow = lastCellWithContentInColumnA() + 1;
    SetBaseLastRow(lastRow)
  }

  var numRows = lastRow - c_START_ROW + 1;
  let range =  sheet.getRange(c_START_ROW, 1, numRows, 7);
  var result = extractRowsFromA1Notation(range.getA1Notation());
  return result
}


function extractRowsFromA1Notation(a1Notation) {
  // Match digits after letters (e.g., A17) and at the end (e.g., G56)
  var match = a1Notation.match(/\D*(\d+)\D*(\d+)/);
  if (match) {
    var startRow = parseInt(match[1]);
    var endRow = parseInt(match[2]);
    return { startRow, endRow };
  } else {
    throw new Error("Invalid A1 notation format: " + a1Notation);
  }
}

function lastCellWithContentInColumnA() {
  ranges = sheet.getRange("A:A");
  var column = ranges.getValues();
  for (var i = column.length - 1; i >= 0; i--) {
    if (column[i][0] !== "") break;
  } ++i
  return i;
}
