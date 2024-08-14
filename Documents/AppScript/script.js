
let startOfMonth = new Date();
let startDate = new Date(2024, 08, 01); //Measured in miliseconds
let endOfMonth = new Date();
let endDate = new Date(2024, 08, 31);
const currentdate = new Date();
const dayOfMonth = new Date().getDate();
const daysInMonth = new Date(currentdate.getFullYear(), currentdate.getMonth()+1, 0).getDate();
const copySheet = 'Copy';
const correctingFactorSheet = 'Correcting Factor';

//In stead of enum, Use maps - https://www.w3schools.com/js/js_maps.asp
//Map email to fullName
//Maybe use both, allowing enum (Dev.CHARLES which contains his email) to be mapped to String (Charles Li)
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

//In stead of enum, Use maps - https://www.w3schools.com/js/js_maps.asp
//Map headerName to index
//Not const
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
  ACTUALTIME: 'Actual time', //Est time - Roll time
  STDTIME: 'Standardised time', //Difference time * correcting factor
  SUMTIME: 'Summed time', //The sum of Actual time
  STDHOURS2DATE: 'Hours to date', //The number of standard hours to date
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
  [HeaderLabels.ACTUALTIME, -1],
  [HeaderLabels.STDTIME, -1],
  [HeaderLabels.SUMTIME, -1],
  [HeaderLabels.STDHOURS2DATE, -1],
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
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
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

function removeLastModified() {
  let sum = 0;

  //Does the 'Completed At' exist
  //function doesColExist(){}
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  let value = sheet.getRange("d1").getValue();
  if(value==='Last Modified') {
    Logger.log('Found "' + value + '"'); 
  } else {
    Logger.log('Did not find. Value is "' + value + '"');
  }

  //Validate and remove excess "Last Modified"
  //function deleteRows(){}
  let data = sheet.getDataRange().getValues();
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows, 3).getValue();
    let cellDate = new Date(cellValue);
    if(!isNaN(cellDate)) {
      if(cellDate <= startDate || celldate >= endDate) {
        //Logger.log('match: ' + cellDate + ' Row: ' + noRows)
        sheet.deleteRow(noRows);
      } 
    } else {
      //Logger.log('invalid date: ' + noRows);
      sheet.deleteRow(noRows);
    }
  }
}

//Remove hard coded indexing of HeaderLabels.COMPLETED
function removeCompletedAt() {
  let sum = 0;

  //Does the 'Completed At' exist
  //function doesColExist(){}
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  let value = sheet.getRange("c1").getValue();
  if(value === HeaderLabels.COMPLETED) {
    //Logger.log('Found "' + value + '"'); 
  } else {
    Logger.log('Did not find. Value is "' + value + '"');
  }

  //Validate and remove excess "Completed At"
  //function deleteRows(){}
  let data = sheet.getDataRange().getValues();
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows, 3).getValue();
    let cellDate = new Date(cellValue);
    if(!isNaN(cellDate)) {
      if(cellDate <= startDate || celldate >= endDate) {
        //Logger.log('match: ' + cellDate + ' Row: ' + noRows)
        sheet.deleteRow(noRows);
      } 
    } else {
      //Logger.log('invalid date: ' + noRows);
      sheet.deleteRow(noRows);
    }
  }
}

//function removeZeroHrs() {}

function separateSharedTasks() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  //Identify column indexes
  //setHeaderIndex();
  let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let devColIndex = HeaderIndex.get(HeaderLabels.DEVELOPER);
  let estTimeColIndex = HeaderIndex.get(HeaderLabels.ESTTIME);
  let rollHrsIndex = HeaderIndex.get(HeaderLabels.ROLLTIME);
  let brandColIndex = HeaderIndex.get(HeaderLabels.BRAND);
  let regionColIndex = HeaderIndex.get(HeaderLabels.REGION);
  let nameColIndex = HeaderIndex.get(HeaderLabels.NAME);
  let techCatIndex = HeaderIndex.get(HeaderLabels.CATEGORY);
  
  //Separate devs
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows,devColIndex).getValue();
    if(cellValue.includes(',')) {
      let splitDevs = cellValue.split(','); //Return the number of devs, not commas.
      let estTimeTemp = sheet.getRange(noRows,estTimeColIndex).getValue();
      let rollHrsTemp = sheet.getRange(noRows,rollHrsIndex).getValue();
      let brandTemp = sheet.getRange(noRows,brandColIndex).getValue();
      let regionTemp = sheet.getRange(noRows,regionColIndex).getValue();
      let nameTemp = sheet.getRange(noRows, nameColIndex).getValue();
      let techCatTemp = sheet.getRange(noRows, techCatIndex).getValue();
      
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
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);

  let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let devColIndex = -1;
  for(let i=0; i<headerRow.length; i++) {
    if(headerRow[i] === 'Dev') {
      Logger.log('devColIndex: '+ (i+1));
      devColIndex = i+1;
    }
  }

  //Format developer names
  let data = sheet.getDataRange().getValues();
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows, devColIndex).getValue();
    switch (cellValue) {
      case MapToDevEmail.BJORN:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.BJORN);
        break;
      case MapToDevEmail.CHARLES:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.CHARLES);
        break;
      case MapToDevEmail.CLYDE:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.CLYDE);
        break;
      case MapToDevEmail.VERNON:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.VERNON);
        break;
      case MapToDevEmail.HITESH:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.HITESH);
        break;
      case MapToDevEmail.RYAN:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.RYAN);
        break;
      case MapToDevEmail.BRENDAN:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.BRENDAN);
        break;
      case MapToDevEmail.CURTIS:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.CURTIS);
        break;
      case MapToDevEmail.DIRK:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.DIRK);
        break;
      case MapToDevEmail.SERGEI:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.SERGEI);
        break;
      case MapToDevEmail.VIJAY:
        sheet.getRange(noRows,devColIndex).setValue(MapToDevName.VIJAY);
        break;
    }
  }
}

//Change to set actual time, then create subsequent function for 'standardisedTime which is actual * correctingfactor
function setActualTime() {
  //add col and do math
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  setHeaderIndex();
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    let estA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ESTTIME)).getA1Notation();
    let rollA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ROLLTIME)).getA1Notation();
    let difference = '=('+estA1+'-'+rollA1+')';
    //sheet.getRange(i,sheet.getLastColumn()).setFormula(difference);
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).setFormula(difference);
  }
}

function setSumOfActualTime() {
  //for each dev, sum the formatted hours
  //Compare to the Number of working days in current month?
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  let d = sheet.getRange(1,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation().replace(/[0-9]/g,'');
  Logger.log('dev index: ' + d);
  let hrs = sheet.getRange(1,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getA1Notation().replace(/[0-9]/g,'');
  Logger.log('hrs index: '+hrs);
  
  //'=SUMIF(' + d + ':' + d + ',' + MapToDevName.name + ',' + hrs + ':' + hrs + ')'
  
  let developerCellNotation;
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    developerCellNotation = sheet.getRange(i,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation();
    Logger.log(developerCellNotation);
    let sumIf = '=SUMIF(' + d + ':' + d + ',' + developerCellNotation + ',' + hrs + ':' + hrs + ')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).setFormula(sumIf);
  }

}

function setMonthToDateHours() {
  setHeaderIndex();
  //Using ((dayOfMonth / daysInMonth)*168)/24
  let formula = '=((' + dayOfMonth + '/' + daysInMonth + ')*168/24)'; //HARDCODE ALERT, BOTH FORMULA AND STD HRS
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  Logger.log('hierso');
  for(let i=2; i<=sheet.getMaxRows(); i++) {
    Logger.log('here');
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.STDHOURS2DATE)).setFormula(formula);
  }
}

function setCorrectingfactor() {
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(copySheet);
  let ithRow = -1;
  let formula = '=IFS(' + /* I need the A1 notation's here... */HeaderIndex.get(HeaderLabels.SUMTIME) + ithRow + '<' + HeaderIndex.get(HeaderLabels.STDHOURS2DATE) + '';
  for(let i=0; i<=sheet.getMaxRows(); i++) {
    //
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.CORRECTFACTOR)).setFormula(formula);
  }
}

function cost() {
  //dur*24*580
}

function main() {
  setHeaderIndex();
  setDate();
  separateSharedTasks();
  formatDev();
}