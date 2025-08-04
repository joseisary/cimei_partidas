/** @OnlyCurrentDoc */
let c_SHIFT_BETWEEN_HEADERS = 4


var spreadsheet = SpreadsheetApp.getActive();
let copyCount = 1;

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
  var sheet     = spreadsheet.getActiveSheet()
  var rangeSrc  = sheet.getRange('17:58'); // 42 rows
  var rows      = rangeSrc.getNumRows();
  var startRow  = 57 ;// where first copy goes
  var offset    = GetCount() * rows;
  var targetRow = startRow + offset + 2; // 2 aditional rows
  var targetRange = sheet.getRange(targetRow, 1, rows, sheet.getMaxColumns());
  
  
  PropertiesService.getScriptProperties().setProperty('status', 'running');
  
  showDialogAndRun()  
  rangeSrc.copyTo(targetRange);
  Collapsing_aux(targetRow);
  Counting()
  PropertiesService.getScriptProperties().setProperty('status', 'done');
}

function DeletedPartida() {
  var sheet = spreadsheet.getSheetByName("COTIZADOR"); 
  var range  = sheet.getRange(57, 1, 43, 7)
  sheet.deleteRows(range.getRow(), range.getNumRows());
};


function Collapsing_aux(baseRow) {

  let rowStart_main        = baseRow + 1
 
  let rowStart_manoDeObra  = baseRow + 3
  let rowEnd_manoDeObra   = getEndRow(rowStart_manoDeObra);

  let rowStart_vOperativos = rowEnd_manoDeObra + c_SHIFT_BETWEEN_HEADERS
  let rowEnd_vOperativos = getEndRow(rowStart_vOperativos);

  let rowStart_materiales = rowEnd_vOperativos + c_SHIFT_BETWEEN_HEADERS
  let rowEnd_materiales = getEndRow(rowStart_materiales)

  let rowStart_otros = rowEnd_materiales + c_SHIFT_BETWEEN_HEADERS
  let rowEnd_otros   =  getEndRow(rowStart_otros)
  
  let rowStart_resument = rowEnd_otros + c_SHIFT_BETWEEN_HEADERS
  let rowEnd_main         = rowStart_resument - 1

  collapseRange(spreadsheet.getRange(rowStart_main        + ':' + rowEnd_main        ), rowStart_main - 1);
  collapseRange(spreadsheet.getRange(rowStart_manoDeObra  + ':' + rowEnd_manoDeObra  ), rowStart_manoDeObra - 1);
  collapseRange(spreadsheet.getRange(rowStart_vOperativos + ':' + rowEnd_vOperativos ), rowStart_vOperativos - 1);
  collapseRange(spreadsheet.getRange(rowStart_materiales  + ':' + rowEnd_materiales  ), rowStart_materiales - 1);
  collapseRange(spreadsheet.getRange(rowStart_otros       + ':' + rowEnd_otros       ), rowStart_otros - 1); 
}
function Collapsing_aux2(baseRow) {

  const spreadsheet = SpreadsheetApp.getActive(); // assumed needed if not global
  const sections = [
    { name: 'main',         offset: 1,   hasContent: false },
    { name: 'manoDeObra',   offset: 3,   hasContent: true  },
    { name: 'vOperativos',  offset: null, hasContent: true },
    { name: 'materiales',   offset: null, hasContent: true },
    { name: 'otros',        offset: null, hasContent: true },
    { name: 'resument',     offset: null, hasContent: false }
  ];

  let rows = {};
  for (let i = 0; i < sections.length; i++) {
    const section = sections[i];

    if (i === 0) {
      // First section 'main'
      rows[section.name + 'Start'] = baseRow + section.offset;
    } else {
      // Next sections start after previous section's end + spacing
      const prevSection = sections[i - 1];
      rows[section.name + 'Start'] = rows[prevSection.name + 'End'] + c_SHIFT_BETWEEN_HEADERS;
    }

    if (section.hasContent) {
      rows[section.name + 'End'] = getEndRow(rows[section.name + 'Start']);
    } else {
      // 'main' and 'resument' rows
      rows[section.name + 'End'] = null;
    }
  }

  // Set 'mainEnd' to the row before 'resumentStart'
  rows['mainEnd'] = rows['resumentStart'] - 1;

  // Collapse all content sections
  ['main', 'manoDeObra', 'vOperativos', 'materiales', 'otros'].forEach(name => {
    const start = rows[name + 'Start'];
    const end = rows[name + 'End'];
    collapseRange(spreadsheet.getRange(`${start}:${end}`), start - 1);
  });
}



function Collapsing(baseRow) {
  // Adjusted collapse ranges, based on original offsets
  
  let shiftStartRow_1 = 1
  let shiftStartRow_2 = 3
  let shiftStartRow_3 = 10
  let shiftStartRow_4 = 20
  let shiftStartRow_5 = 30

  let shiftEndRow_1 = 39
  let shiftEndRow_2 = 6
  let shiftEndRow_3 = 16
  let shiftEndRow_4 = 26
  let shiftEndRow_5 = 36
  
  let startRow_1 = baseRow + shiftStartRow_1
  let startRow_2 = baseRow + shiftStartRow_2
  let startRow_3 = baseRow + shiftStartRow_3
  let startRow_4 = baseRow + shiftStartRow_4
  let startRow_5 = baseRow + shiftStartRow_5

  let endRow_1 = baseRow + shiftEndRow_1
  let endRow_2 = baseRow + shiftEndRow_2
  let endRow_3 = baseRow + shiftEndRow_3
  let endRow_4 = baseRow + shiftEndRow_4
  let endRow_5 = baseRow + shiftEndRow_5


  console.log("the startRow_1 is at " + startRow_1)
  console.log("the endRow_1 is at " + endRow_1)

  console.log("the startRow_2 is at " + startRow_2)
  console.log("the endRow_2 is at " + endRow_2)
  
  console.log("the startRow_3 is at " + startRow_3)
  console.log("the endRow_3 is at " + endRow_3)

  console.log("the startRow_4 is at " + startRow_4)
  console.log("the endRow_4 is at " + endRow_4)

  console.log("the startRow_5 is at " + startRow_5)
  console.log("the endRow_5 is at " + endRow_5)

  collapseRange(spreadsheet.getRange(startRow_1  + ':' + endRow_1 ), startRow_1 - 1);
  collapseRange(spreadsheet.getRange(startRow_2  + ':' + endRow_2 ), startRow_2 - 1);
  collapseRange(spreadsheet.getRange(startRow_3  + ':' + endRow_3 ), startRow_3 - 1);
  collapseRange(spreadsheet.getRange(startRow_4  + ':' + endRow_4 ), startRow_4 - 1);
  collapseRange(spreadsheet.getRange(startRow_5  + ':' + endRow_5 ), startRow_5 - 1); 
}


function collapseRange(range, startRow) {
  range.shiftRowGroupDepth(1);

  sheet = spreadsheet.getActiveSheet()
  group = sheet.getRowGroup(startRow, 1)
  group.collapse()
}

function MoveButton(sheet, row, col) {
  var drawings = sheet.getDrawings();

  for (var i = 0; i < drawings.length; i++) {
    if (drawings[i].getOnAction() === "Entry") { // Check if this is the "+" button
      drawings[i].setPosition(sheet.getRange(row, col), 0, 0);
      break;
    }
  }
}

var partidasGenerated_range = spreadsheet.getActiveSheet().getRange('A1');
function GetCount(){
  var count = partidasGenerated_range.getValue()
  return count === "" ? 0 : count;
}

function Counting(){
  var copyCount = partidasGenerated_range.getValue();
  copyCount = copyCount === "" ? 0 : copyCount;
  partidasGenerated_range.setValue(copyCount + 1);
}

function ResetCounting(){// attached by a button sheet
  partidasGenerated_range.setValue(0);
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

function getEndRow(startRow){
  let i = 0
  for( i = startRow + 1;  !isRowEmpty(i); i++);
  i--;
  return i;

}

function getEndRow_v1(startRow) {
  if (startRow < 1) throw new Error("startRow must be >= 1");
  let row = startRow + 1;
  while (!isRowEmpty(row)) row++;
  return row - 1;
}

function isRowEmpty(rowNumber) {
   var c_LAST_COLUMN = 6
   
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowValues = sheet.getRange(rowNumber, 1, 1, c_LAST_COLUMN).getValues()[0];
  // Check if every cell is empty
  return rowValues.every(function(cell) {
    return cell === "" || cell === null;
  });
}


