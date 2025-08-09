let c_CELL_PERSISTENCE = "I"

var countingPartidas      = sheet.getRange(c_CELL_PERSISTENCE + '1');
var baseLastRowSrcPersist = sheet.getRange(c_CELL_PERSISTENCE + '2');
var lastStartRowPersist   = sheet.getRange(c_CELL_PERSISTENCE + '3');
var lastRowsPersist       = sheet.getRange(c_CELL_PERSISTENCE + '4');


function IncreaseCountByOne(){
  var copyCount = countingPartidas.getValue();
  copyCount = copyCount === "" ? 0 : copyCount;
  countingPartidas.setValue(copyCount + 1);
}
function DecreaseCountByOne(){
  var copyCount = countingPartidas.getValue();
  copyCount = copyCount === "" ? 0 : copyCount;
  copyCount = copyCount === 0 ? 0 : copyCount - 1;
  countingPartidas.setValue(copyCount);
}
function ReadCount(){
  var count = countingPartidas.getValue()
  return count === "" ? 0 : count;
}
function ResetCounting(){// attached by a button sheet
  countingPartidas.setValue(0);
}

function GetLastStartRow(){
  var val = lastStartRowPersist.getValue()
  return val === "" ? 0 : val;
}
function SetLastStartRow(value){
  lastStartRowPersist.setValue(value);
}

function GetLastRows(){
  var val = lastRowsPersist.getValue()
  return val === "" ? 0 : val;
}

function SetLastRows(value){
  lastRowsPersist.setValue(value);
}

function GetBaseLastRow(){
  var val = baseLastRowSrcPersist.getValue()
  return val === "" ? 0 : val;
}
function SetBaseLastRow(value){
  baseLastRowSrcPersist.setValue(value);
}
