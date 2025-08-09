/** @OnlyCurrentDoc */
let c_SHIFT_BETWEEN_HEADERS  = 4
let c_START_ROW              = 17
let c_SHIFT_OVER_PARTIDA     = 1
let c_LAST_COLUMN_WITH_DATA  = 7

var spreadsheet = SpreadsheetApp.getActive();
var sheet = spreadsheet.getSheetByName("COTIZADOR")

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛠 Custom Tools")
    .addItem("Show Reset Button", "ShowResetSidebar")
    .addToUi();
}

function AddPartida(){
  Entry()
}

function Entry() {
  ShowResetSidebar()
  GenPartida()
}

function GenPartida() {
  var val = getRangeSrc()
  var rangeSrc  = sheet.getRange(val.startRow + ':' + val.endRow );
  var rows      = rangeSrc.getNumRows();
  var startRow  = rows + c_START_ROW + c_SHIFT_OVER_PARTIDA
  var offset    = ReadCount() * ( rows + c_SHIFT_OVER_PARTIDA); 
  var targetRow = startRow + offset;
  var targetRange = sheet.getRange(targetRow, 1, rows, sheet.getMaxColumns());
  
  PropertiesService.getScriptProperties().setProperty('status', 'running');
  
  showDialogAndRun()  
  rangeSrc.copyTo(targetRange);
  Collapsing(targetRow);
  IncreaseCountByOne()  

  /* update globals */
  SetLastStartRow(targetRow)
  SetLastRows(rows)
  PropertiesService.getScriptProperties().setProperty('status', 'done');
}

function DeletedPartida() {
  if(ReadCount() == 0 ){
    return
  }
  var targetRow = GetLastStartRow()
  var rows      = GetLastRows()
  sheet.deleteRows(targetRow, rows);
  targetRow = targetRow  - (rows + c_SHIFT_OVER_PARTIDA)
  SetLastStartRow(targetRow)
  DecreaseCountByOne()
};

function Collapsing(baseRow) {
  let sections = [];

  let rowStart_main = baseRow + 1;
  let rowStart = baseRow + 3;

  const sectionNames = [
    "manoDeObra",
    "vOperativos",
    "materiales",
    "otros"
  ];

  for (let name of sectionNames) {
    let rowEnd = getEndRow(rowStart);
    sections.push({ name, rowStart, rowEnd });
    rowStart = rowEnd + c_SHIFT_BETWEEN_HEADERS;
  }

  let rowStart_resument = rowStart;
  let rowEnd_main = rowStart_resument - 1;

  // Collapse main section
  collapseRange(spreadsheet.getRange(`${rowStart_main}:${rowEnd_main}`), rowStart_main - 1);

  // Collapse all content sections
  for (let section of sections) {
    collapseRange(
      spreadsheet.getRange(`${section.rowStart}:${section.rowEnd}`),
      section.rowStart - 1
    );
  }
}

function showDialogAndRun() {
  const html = HtmlService.createHtmlOutputFromFile('dialog')
    .setWidth(300)
    .setHeight(100);
  SpreadsheetApp.getUi().showModelessDialog(html, "Processing...");
}

function ShowResetSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('buttonAdd')
      .setTitle('Reset Counter');
  SpreadsheetApp.getUi().showSidebar(html);
}

// This is polled by HTML
function getStatus() {
  return PropertiesService.getScriptProperties().getProperty('status');
}

function searching() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:A').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.UP).activate();

};

function getLastCell() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var lastCell = sheet.getRange(lastRow, lastColumn);
  
  console.log("Last cell address: " + lastCell.getA1Notation());
  return lastCell;
}



