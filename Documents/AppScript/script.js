
let startOfMonth = new Date();
let startDate = new Date(2024, 7, 01); //Measured in miliseconds
let endOfMonth = new Date();
let endDate = new Date(2024, 7, 31);
const currentdate = new Date();
const dayOfMonth = new Date().getDate();
const daysInMonth = new Date(currentdate.getFullYear(), currentdate.getMonth()+1, 0).getDate();
const dataSheet = 'Copy';
const correctingFactorSheet = 'Correcting Factor';
const costFactor = 580*24; //Note that the costFactor needs to be *24 to convert to int

//Configurable enums / Maps
//Further require enums for formulas * all date values above
const MapToDevEmail = {
  CHARLES: 'charles.li@velosure.com.au', //and add a value 'Charles Li'
  CLYDE: 'clyde@twothreebird.com',
  BJORN: 'bjorn@twothreebird.com',
  VERNON: 'vernon@twothreebird.com',
  HITESH: 'hitesh@twothreebird.com',
  RYAN: 'ryan@twothreebird.com',
  CURTIS: 'curtis@twothreebird.com',
  DIRK: 'dirk@twothreebird.com',
  BRENDAN: 'brendan@twothreebird.com',
  SERGEI: 'sergei@twothreebird.com',
  VIJAY: 'vijay@twothreebird.com',
};

const MapToDevName = {
  CHARLES: 'Charles Li', //and add a value 'Charles Li'
  CLYDE: 'Clyde Cyster',
  BJORN: 'Björn Theart',
  VERNON: 'Vernon Grant',
  HITESH: 'Hitesh Maity',
  RYAN: 'Ryan Peel',
  CURTIS: 'Curtis Page',
  DIRK: 'Dirk Dircksen',
  BRENDAN: 'Brendan van der Meulen',
  SERGEI: 'Sergei Pringiers',
  VIJAY: 'Vijay Kumar',
}

const HeaderLabels = {
  CREATED: 'Created At',
  COMPLETED: 'Completed At',
  CORRECTFACTOR: 'Correcting factor',
  MODIFIED: 'Last Modified',
  NAME: 'Name',
  PROGRESS: 'Tech Progress',
  BRAND: 'Brand',
  DEVELOPER: 'Dev',
  CATEGORY: 'Tech Category',
  REGION: 'Region',
  ESTTIME: 'Estimated time', //Time logged on Asana
  ROLLTIME: 'Rollover time', //Time spent in previous reporting period
  DIFFTIME: 'Difference time', //NA
  ACTUALTIME: 'Actual time', //Est time - Roll time
  STDTIME: 'Standardised time', //Difference time * correcting factor
  SUMTIME: 'Summed time', //The sum of Actual time
  MTDHRS: 'Hours to date', //The number of standard hours to date
  COST: 'Cost'
};

let HeaderIndex = new Map([
  [HeaderLabels.CREATED, -1],
  [HeaderLabels.COMPLETED, -1],
  [HeaderLabels.CORRECTFACTOR, -1],
  [HeaderLabels.MODIFIED, -1],
  [HeaderLabels.NAME, -1],
  [HeaderLabels.PROGRESS, -1],
  [HeaderLabels.BRAND, -1],
  [HeaderLabels.DEVELOPER, -1],
  [HeaderLabels.CATEGORY, -1],
  [HeaderLabels.REGION, -1],
  [HeaderLabels.ESTTIME, -1],
  [HeaderLabels.ROLLTIME, -1],
  [HeaderLabels.DIFFTIME, -1],
  [HeaderLabels.ACTUALTIME, -1],
  [HeaderLabels.STDTIME, -1],
  [HeaderLabels.SUMTIME, -1],
  [HeaderLabels.MTDHRS, -1],
  [HeaderLabels.COST, -1]
]);

function setDate() {
  startOfMonth.setDate(1);
  startOfMonth.setHours(0,0,0,0);
  endOfMonth.setMonth(startOfMonth.getMonth()+1);
  endOfMonth.setDate(0);
  endOfMonth.setHours(0,0,0,0);
}

function setHeaderIndex() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  //If some headers are not found, create column
  for(let i=0; i<headers.length; i++) {
    HeaderIndex.forEach(function(value, key) {
      if(headers[i] === key) {
        HeaderIndex.set(key, i+1);
      }
    })
  }
  
  HeaderIndex.forEach(function(value,key) {
    //check if mapped to -1
    if(value < 0) {
      if(sheet.getLastColumn() == sheet.getMaxColumns()) {
        sheet.insertColumnAfter(sheet.getLastColumn());
      }
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(key);
      HeaderIndex.set(key, sheet.getRange(1, sheet.getLastColumn()));
    }
  })
}

function clearLastModified() {
  /* 
  Apply filter
  Sort by last modified
  Clear contents
  */
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  let filter = sheet.getFilter();
  if(filter) {
    filter.remove();
  }
  range.createFilter();
  sheet.sort(HeaderIndex.get(HeaderLabels.MODIFIED), false);
  sheet.getFilter().remove();
  
  let startingRow = -1;
  let lastModified;
  for(let i=2; i<=sheet.getLastRow(); i++) { //.getLastRow() returns the last row that contains content
    lastModified = new Date(sheet.getRange(i, HeaderIndex.get(HeaderLabels.MODIFIED)).getValue());
      startingRow = i;
    if(lastModified < startDate) {
      break;
    }
  }
  
  Logger.log(startingRow);
  let clearColumnFrom = 1;
  let numberOfRows = sheet.getMaxRows()-startingRow;
  let lastColumn = sheet.getLastColumn();
  let clearRange = sheet.getRange(startingRow,1,numberOfRows,lastColumn);
  clearRange.clearContent();
}

//Review filter and consider blank rows
/*
The problem is now, if I clear, the cleared range is incorrect.
Maybe arrang inverse, the start from where the completed at is not blank.
THIS WILL WORK! Because the filtered values only apply to cells which contains content. if not content, it's appended below.
*/
function clearCompletedAt() {
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  let filter = sheet.getFilter();
  if(filter) {
    filter.remove();
  }
  filter = range.createFilter();
  //filter.setColumnFilterCriteria(HeaderIndex.get(HeaderLabels.COMPLETED), SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build());
  sheet.getFilter().sort(HeaderIndex.get(HeaderLabels.COMPLETED), true);
  
  //sheet.getFilter().remove();

  /*let startingRow = -1;
  let completedAt;
  for(let i=2; i<=sheet.getLastRow(); i++) { //.getLastRow() returns the last row that contains content
    completedAt = new Date(sheet.getRange(i, HeaderIndex.get(HeaderLabels.COMPLETED)).getValue());
      startingRow = i;
    if(completedAt < startDate) {
      break;
    }
  }
  
  Logger.log(startingRow);
  let clearColumnFrom = 1;
  let numberOfRows = sheet.getMaxRows()-startingRow;
  let lastColumn = sheet.getLastColumn();
  let clearRange = sheet.getRange(startingRow,1,numberOfRows,lastColumn);
  clearRange.clearContent();*/
  
  
}

/*
function removeZeroHrs() {
  //If actual hrs is zero, delete the row
  setHeaderIndex();
  setActualTime();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=sheet.getMaxRows(); i>1; i--) {
    let range = sheet.getRange(i,1,1,sheet.getLastColumn());
    if(sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getValue() == 0 || sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getValue() == isNaN) {
      range.clearContent();
    }
  }
}
*/

function separateSharedTasks() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //Identify column indexes
  //setHeaderIndex();
  let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  //Separate devs
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows,devColIndex).getValue();
    if(cellValue.includes(',')) {
      let splitDevs = cellValue.split(','); //Return the number of devs, not commas.
      let estTimeTemp = sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.ESTTIME)).getValue();
      let rollHrsTemp = sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.ROLLTIME)).getValue();
      let brandTemp = sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.BRAND)).getValue();
      let regionTemp = sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.REGION)).getValue();
      let nameTemp = sheet.getRange(noRows, HeaderIndex.get(HeaderLabels.NAME)).getValue();
      let techCatTemp = sheet.getRange(noRows, HeaderIndex.get(HeaderLabels.CATEGORY)).getValue();
      
      for(let i=0; i<cellValue.split(',').length-1; i++) {
        sheet.insertRowAfter(noRows);
        sheet.getRange(noRows+1,devColIndex).setValue(splitDevs[i+1]).trimWhitespace();
        sheet.getRange(noRows+1,nameColIndex).setValue(nameTemp);
        sheet.getRange(noRows+1,regionColIndex).setValue(regionTemp);
        sheet.getRange(noRows+1,brandColIndex).setValue(brandTemp);
        sheet.getRange(noRows+1,rollHrsIndex).setValue(rollHrsTemp);
        sheet.getRange(noRows+1,estTimeColIndex).setValue(estTimeTemp);
        sheet.getRange(noRows+1,techCatIndex).setValue(techCatTemp);
      }
      sheet.getRange(noRows,devColIndex).setValue(splitDevs[0]);
    }
  }
}

function formatDev() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //Format developer names
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows, HeaderIndex.get(HeaderLabels.DEVELOPER)).getValue();
    switch (cellValue) {
      case MapToDevEmail.BJORN:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.BJORN);
        break;
      case MapToDevEmail.CHARLES:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.CHARLES);
        break;
      case MapToDevEmail.CLYDE:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.CLYDE);
        break;
      case MapToDevEmail.VERNON:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.VERNON);
        break;
      case MapToDevEmail.HITESH:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.HITESH);
        break;
      case MapToDevEmail.RYAN:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.RYAN);
        break;
      case MapToDevEmail.BRENDAN:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.BRENDAN);
        break;
      case MapToDevEmail.CURTIS:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.CURTIS);
        break;
      case MapToDevEmail.DIRK:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.DIRK);
        break;
      case MapToDevEmail.SERGEI:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.SERGEI);
        break;
      case MapToDevEmail.VIJAY:
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(MapToDevName.VIJAY);
        break;
    }
  }
}

function setActualTime() {
  //add col and do math
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //setHeaderIndex();
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    let estA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ESTTIME)).getA1Notation();
    let rollA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ROLLTIME)).getA1Notation();
    let difference = '=('+estA1+'-'+rollA1+')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).setFormula(difference);
  }
}

function setSumOfActualTime() {
  //for each dev, sum the formatted hours
  //Compare to the Number of working days in current month?
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let d = sheet.getRange(1,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation().replace(/[0-9]/g,'');
  let hrs = sheet.getRange(1,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getA1Notation().replace(/[0-9]/g,'');
  
  let developerCellNotation;
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    developerCellNotation = sheet.getRange(i,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation();
    let sumIf = '=SUMIF(' + d + ':' + d + ',' + developerCellNotation + ',' + hrs + ':' + hrs + ')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).setFormula(sumIf);
  }

}

function setMonthToDateHours() {
  //setHeaderIndex();
  //Using ((dayOfMonth / daysInMonth)*168)/24
  let formula = '=((' + dayOfMonth + '/' + daysInMonth + ')*168/24)'; //HARDCODE ALERT, BOTH FORMULA AND STD HRS
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).setFormula(formula);
  }
}

function setCorrectingfactor() {
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  /* 
  if: SUMTIME<(MTDHRS/24)
  then: (MTDHRS/24)/SUMTIME
  else if: SUMTIME>(MTDHRS/24)
  then: 1/( (MTDHRS/24)/SUMTIME )
  */
  let i = -1;
  for(i=2; i<=sheet.getMaxRows(); i++) {
    let condition1 = '('+sheet.getRange(i, HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+'<'+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+')';
    let then1 = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+'/'+sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+')';
    let condition2 = '('+sheet.getRange(i, HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+'>'+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+')';
    let then2 = '1/('+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+'/'+sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+')';
    let correctingFactorFormula = '=IFS('+condition1+','+then1+','+condition2+','+then2+')';
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.CORRECTFACTOR)).setFormula(correctingFactorFormula);
  }
}

function setStandardisedHours() {
  //correctingFactor * actualHours
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    let stdHrsFormula = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getA1Notation()+'*'+sheet.getRange(i,HeaderIndex.get(HeaderLabels.CORRECTFACTOR)).getA1Notation()+')';
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.STDTIME)).setFormula(stdHrsFormula);
  }
}

function setCost() {
  //dur*24*580
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    let costFormula = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.STDTIME)).getA1Notation()+'*'+costFactor+')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.COST)).setFormula(costFormula);
  }
}

function main() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  setDate();
  setHeaderIndex();
  clearLastModified();
  clearCompletedAt();
  separateSharedTasks();
  formatDev();
  setActualTime();
  setSumOfActualTime();
  setMonthToDateHours();
  setCorrectingfactor();
  setStandardisedHours();
  setCost();
}