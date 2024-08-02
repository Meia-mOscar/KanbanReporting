/**Remove bloat data
 * Match developer; SUM "Est. Time"; Conditions[{Completed }{""}{current month}]
 * Delete lines not matching these conditions 
 * The active sheet is tab 0
*/

let startDate = new Date(2024, 06, 30); //Measured in miliseconds
let endDate = new Date(2024, 08, 01);
const sheetName = 'Copy';

const Devs = {
  CHARLES: 'charles.li@velosure.com.au',
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

//Add switch in separateDevs() for header index
//Or foreach() enum compares to headers
const Headers = {
  CREATED: 'Created At',
  COMPLETED: 'Completed At',
  MODIFIED: 'Last Modified',
  NAME: 'Name',
  PROGRESS: 'Tech Progress',
  BRAND: 'Brand',
  DEVELOPER: 'Dev',
  CATEGORY: 'Tech Category',
  REGION: 'Region',
  ESTTIME: 'Estimated time',
  ROLLTIME: 'Rollover time',
  STDTIME: 'Standardised time',
  COST: 'Cost'
}

function removeLastModified() {
  let sum = 0;

  //Does the 'Completed At' exist
  //function doesColExist(){}
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy');
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

function removeCompletedAt() {
  let sum = 0;

  //Does the 'Completed At' exist
  //function doesColExist(){}
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy');
  let value = sheet.getRange("c1").getValue();
  if(value==='Completed At') {
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

function separateDev() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy');
  if(!sheet) {
    Logger.log('Sheet not found: ' + sheetName)
  }
  //Identify column indexes
  let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let devColIndex = -1;
  let estTimeColIndex = -1;
  let rollHrsIndex = -1;
  let brandColIndex = -1;
  let regionColIndex = -1;
  let nameColIndex = -1;
  for(let i=0; i<headerRow.length; i++) {
    if(headerRow[i] === 'Dev') {
      Logger.log('devColIndex: '+ (i+1));
      devColIndex = i+1;
    } else if(headerRow[i] === 'Estimated time') {
      Logger.log('Estimated time: ' + (i+1));
      estTimeColIndex = i+1;
    } else if(headerRow[i] === 'Brand') {
      Logger.log('Brand' + (i+1));
      brandColIndex = i+1;
    } else if(headerRow[i] === 'Region') {
      Logger.log('Region' + (i+1));
      regionColIndex = i+1;
    } else if(headerRow[i] === 'Name') {
      Logger.log('Name: ' + (i+1));
      nameColIndex = i+1;
    } else if(headerRow[i] === 'Rollover time') {
      rollHrsIndex = i+1;
    }
  }
  //Separate devs
  let cellValue = '';
  for(let noRows = sheet.getMaxRows(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows,devColIndex).getValue();
    if(cellValue.includes(',')) {
      Logger.log('comma: ' + noRows + ' devs: ' + cellValue.split(',').length);
      let splitDevs = cellValue.split(','); //Return the number of devs, not commas.
      let estTimeTemp = sheet.getRange(noRows,estTimeColIndex).getValue();
      let rollHrsTemp = sheet.getRange(noRows,rollHrsIndex).getValue();
      let brandTemp = sheet.getRange(noRows,brandColIndex).getValue();
      let regionTemp = sheet.getRange(noRows,regionColIndex).getValue();
      let nameTemp = sheet.getRange(noRows, nameColIndex).getValue();
      for(let i=0; i<cellValue.split(',').length-1; i++) {
        sheet.insertRowAfter(noRows);
        sheet.getRange(noRows+1,devColIndex).setValue(splitDevs[i+1]).trimWhitespace();
        sheet.getRange(noRows+1,nameColIndex).setValue(nameTemp);
        sheet.getRange(noRows+1,regionColIndex).setValue(regionTemp);
        sheet.getRange(noRows+1,brandColIndex).setValue(brandTemp);
        sheet.getRange(noRows+1,rollHrsIndex).setValue(rollHrsTemp);
        sheet.getRange(noRows+1,estTimeColIndex).setValue(estTimeTemp);
      }
      sheet.getRange(noRows,devColIndex).setValue(splitDevs[0]);
    }
  }
}

function formatDev() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy');
  if(!sheet) {
    Logger.log('Sheet not found: ' + sheetName)
  }

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
      case Devs.BJORN:
        sheet.getRange(noRows,devColIndex).setValue('Björn Theart');
        break;
      case Devs.CHARLES:
        sheet.getRange(noRows,devColIndex).setValue('Charles Li');
        break;
      case Devs.CLYDE:
        sheet.getRange(noRows,devColIndex).setValue('Clyde Cyster');
        break;
      case Devs.VERNON:
        sheet.getRange(noRows,devColIndex).setValue('Vernon Grant');
        break;
      case Devs.HITESH:
        sheet.getRange(noRows,devColIndex).setValue('Hitesh Maity');
        break;
      case Devs.RYAN:
        sheet.getRange(noRows,devColIndex).setValue('Ryan Peel');
        break;
      case Devs.BRENDAN:
        sheet.getRange(noRows,devColIndex).setValue('Brendan van der Meulen');
        break;
      case Devs.CURTIS:
        sheet.getRange(noRows,devColIndex).setValue('Curtis Page');
        break;
      case Devs.DIRK:
        sheet.getRange(noRows,devColIndex).setValue('Dirk Dircksen');
        break;
      case Devs.SERGEI:
        sheet.getRange(noRows,devColIndex).setValue('Sergei Pringiers');
        break;
      case Devs.VIJAY:
        sheet.getRange(noRows,devColIndex).setValue('Vijay Kumar');
        break;
    }
  }
}

function formatTime() {
  //add col and do math
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if(!sheet) {
    Logger.log('"' + sheetName + '" not found ');
    return;
  }

  //let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  //Find Est time and Rollover time
  let cellValue = '';
  let estIndex = -1;
  let rollIndex = -1;
  let headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  for(let i=1; i<headerRow.length; i++) {
    cellValue = sheet.getRange(1,i).getValue();
    if(cellValue == Headers.ESTTIME) {
      estIndex = i;
      Logger.log('est time: ' + estIndex + 'cellValue: ' + cellValue);
    } else if (cellValue == Headers.ROLLTIME) {
      rollIndex = i;
      Logger.log('roll time: ' + rollIndex + 'cellValue: ' + cellValue);
    }
  }
  //Add column and do match
  sheet.insertColumnAfter(sheet.getLastColumn());
  sheet.getRange(1,sheet.getLastColumn()+1).setValue(Headers.STDTIME);
  //Difference calc not working
  for(let i=2; i<sheet.getMaxRows(); i++) {
    sheet.getRange(i,sheet.getLastColumn()).setValue(sheet.getRange(i,estIndex).getValue() - sheet.getRange(i,rollIndex).getValue());
  }
}

function cost() {
  //dur*24*580
}

function correctingFactor() {
  //for each dev, sum the formatted hours
  //Compare to the Number of working days in current month?
}

/**Apply correcting factor
 * 
 * 
 * 
*/

function correctTime() {

}