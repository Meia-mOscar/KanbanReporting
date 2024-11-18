/* Actions, to do:
 * 1. Refactor main & foo's  to bounce 'sheet' from one to another
 * 2. Re-work the search functions - binary / bubble sort / something faster
*/

const dataSheet = 'Copy';

let Configs = {
  STARTDATE: new Date(),
  ENDDATE: new Date(),
  DAYOFMONTH: -1,
  DAYSINMONTH: -1,
  COSTFACTOR: 0
}

const MapEmailToName = new Map([
  ['charles.li@velosure.com.au', 'Charles Li'],
  ['clyde@twothreebird.com', 'Clyde Cyster'],
  ['bjorn@twothreebird.com', 'Björn Theart'],
  ['vernon@twothreebird.com','Vernon Grant'],
  ['hitesh@twothreebird.com', 'Hitesh Maity'],
  ['ryan@twothreebird.com', 'Ryan Peel'],
  ['curtis@twothreebird.com', 'Curtis Page'],
  ['dirk@twothreebird.com', 'Dirk Dircksen'],
  ['brendan@twothreebird.com', 'Brendan van der Meulen'],
  ['sergei@twothreebird.com', 'Sergei Pringiers'],
  ['vijay@twothreebird.com', 'Vijay Kumar'],
  ['lara.ferroni@twothreebird.com','Lara Ferroni'],
  ['lara@project529.com','Lara Ferroni'],
  ['brett.field@twothreebird.com','Brett Field']
]);

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

function setConfigs(isEndOfMonth) {  
  if(isEndOfMonth == true) {
    Configs.DAYSINMONTH = new Date(Configs.STARTDATE.getFullYear(), Configs.STARTDATE.getMonth(), 0).getDate(); /*EOM */
  } else {
    Configs.DAYSINMONTH = new Date(Configs.STARTDATE.getFullYear(), Configs.STARTDATE.getMonth()+1, 0).getDate();
  }
  Logger.log(Configs.DAYSINMONTH);

  if(isEndOfMonth == true) {
    Configs.STARTDATE = new Date(2024,9,1); /*EOM*/
  } else {
    Configs.STARTDATE.setDate(1);
  }
  Configs.STARTDATE.setHours(0,0,0,0);
  Logger.log(Configs.STARTDATE);

  Configs.ENDDATE.setMonth(Configs.STARTDATE.getMonth()+1);
  Configs.ENDDATE.setDate(0);
  Configs.ENDDATE.setHours(0,0,0,0);
  Logger.log("End date " + Configs.ENDDATE);

  if(isEndOfMonth == true) {
    Configs.DAYOFMONTH = Configs.DAYSINMONTH;
  } else {
    Configs.DAYOFMONTH = new Date().getDate();
  }
  Logger.log("Day of month " + Configs.DAYOFMONTH);

  Configs.COSTFACTOR = 580*24; //Duration must be *24 to convert to int. This is done here.
}

function setHeaderIndex() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
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
  setHeaderIndex();
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
    if(lastModified < Configs.STARTDATE) {
      break;
    }
  }
  
  //Logger.log(startingRow);
  let clearColumnFrom = 1;
  let numberOfRows = sheet.getMaxRows()-startingRow;
  let lastColumn = sheet.getLastColumn();
  let clearRange = sheet.getRange(startingRow,clearColumnFrom,numberOfRows,lastColumn);
  clearRange.clearContent();
}

function clearCompletedAt() {
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let filter = sheet.getFilter();
  if(filter) {
    filter.remove();
  }
  let range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  filter = range.createFilter();
  sheet.getFilter().sort(HeaderIndex.get(HeaderLabels.COMPLETED), false);
  sheet.getFilter().remove();

  let completedAtValues = sheet.getRange(1,HeaderIndex.get(HeaderLabels.COMPLETED), sheet.getLastRow()).getValues();
  let indexOfBlank = completedAtValues.findIndex(row => row[0] === "");
  let lastRow = indexOfBlank+1;
  //Logger.log('indexOfBlank: ' + indexOfBlank);

  let completedAt;
  let startingRow = -1;
  for(let i=2; i<=sheet.getLastRow(); i++) { /*.getLastRow() returns the last row that contains content (considering all cols)*/
    completedAt = new Date(sheet.getRange(i, HeaderIndex.get(HeaderLabels.COMPLETED)).getValue());
    if(completedAt < Configs.STARTDATE) {
      startingRow = i;
      break;
    }
  }
  if(startingRow>0) {
    let clearColumnFrom = 1;
    let numberOfRows = lastRow-startingRow;
    let lastColumn = sheet.getLastColumn();
    let clearRange = sheet.getRange(startingRow,clearColumnFrom,numberOfRows,lastColumn);
    //clearRange.clearContent();
    sheet.deleteRows(startingRow,numberOfRows);
  }
}

function clearZeroEst() {
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let filter = sheet.getFilter();
  if(filter) {
    filter.remove();
  }
  let range = sheet.getRange(1,1,sheet.getLastRow(),sheet.getLastColumn());
  filter = range.createFilter();
  sheet.getFilter().sort(HeaderIndex.get(HeaderLabels.ESTTIME), false);
  sheet.getFilter().remove();
  
  let estTimeValues = sheet.getRange(1,HeaderIndex.get(HeaderLabels.ESTTIME), sheet.getLastRow()).getValues();
  let startingRow = estTimeValues.findIndex(row => row[0] === "")+1;
  let lastRow = sheet.getLastRow();

  if(startingRow>0) {
    let numberOfRows = lastRow-startingRow;
    let lastColumn = sheet.getLastColumn();
    let clearColumnFrom = 1;

    let clearRange = sheet.getRange(startingRow,clearColumnFrom,numberOfRows,lastColumn);
    clearRange.clearContent();
    //sheet.deleteRows(startingRow,numberOfRows);
  }
}

function separateSharedTasks() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //Identify column indexes
  setHeaderIndex();
  let cellValue = '';
  for(let x = sheet.getLastRow(); x>0; x--) {
    cellValue = sheet.getRange(x,HeaderIndex.get(HeaderLabels.DEVELOPER)).getValue();
    if(cellValue.includes(',')) {
      let splitDevs = cellValue.split(','); //Return the number of devs, not commas.      
      for(let y=0; y<cellValue.split(',').length-1; y++) {
        sheet.insertRowAfter(x);
        //source_range.copy(target_range);
        sheet.getRange(x,1,1,sheet.getLastColumn()).copyTo(sheet.getRange(x+1,1,1,sheet.getLastColumn()));
        sheet.getRange(x+1,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(splitDevs[y+1]).trimWhitespace();
      }
      sheet.getRange(x,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(splitDevs[0]);
    }
  }
}

function formatDev() {
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //Format developer names
  let cellValue = '';
  for(let noRows = sheet.getLastRow(); noRows>0; noRows--) {
    cellValue = sheet.getRange(noRows, HeaderIndex.get(HeaderLabels.DEVELOPER)).getValue();

    MapEmailToName.forEach(function(value, key){
      if(cellValue === key) {
        sheet.getRange(noRows,HeaderIndex.get(HeaderLabels.DEVELOPER)).setValue(value);
      }
    })
  }
}

function setActualTime() {
  //add col and do math
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //setHeaderIndex();
  for(let i=2; i<=sheet.getLastRow(); i++) {
    let estA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ESTTIME)).getA1Notation();
    let rollA1 = sheet.getRange(i,HeaderIndex.get(HeaderLabels.ROLLTIME)).getA1Notation();
    let difference = '=('+estA1+'-'+rollA1+')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).setFormula(difference);
  }
}

function setSumOfActualTime() {
  //for each dev, sum the formatted hours
  //Compare to the Number of working days in current month?
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  let d = sheet.getRange(1,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation().replace(/[0-9]/g,'');
  let hrs = sheet.getRange(1,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getA1Notation().replace(/[0-9]/g,'');
  
  let developerCellNotation;
  for(let i=2; i<=sheet.getLastRow(); i++) {
    developerCellNotation = sheet.getRange(i,HeaderIndex.get(HeaderLabels.DEVELOPER)).getA1Notation();
    let sumIf = '=SUMIF(' + d + ':' + d + ',' + developerCellNotation + ',' + hrs + ':' + hrs + ')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).setFormula(sumIf);
  }

}

function setMonthToDateHours() {
  setConfigs();
  setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  Logger.log(sheet.getRange(1,HeaderIndex.get(HeaderLabels.MTDHRS),1).getA1Notation());
  let range = sheet.getRange(2,HeaderIndex.get(HeaderLabels.MTDHRS),sheet.getLastRow()); //get_range: row, col, number_rows
  let formula = '=(('+Configs.DAYOFMONTH+'/'+Configs.DAYSINMONTH+')*168/24)';
  range.setFormula(formula);
}

function setCorrectingfactor() {
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  /* 
  if: SUMTIME<(MTDHRS/24)
  then: (MTDHRS/24)/SUMTIME
  else if: SUMTIME>(MTDHRS/24)
  then: 1/(SUMTIME/(MTDHRS/24))
  */
  let i = -1;
  for(i=2; i<=sheet.getLastRow(); i++) {
    let condition1 = '('+sheet.getRange(i, HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+'<'+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+')';
    let then1 = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+'/'+sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+')';
    let condition2 = '('+sheet.getRange(i, HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+'>'+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+')';
    let then2 = '1/('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.SUMTIME)).getA1Notation()+'/'+sheet.getRange(i, HeaderIndex.get(HeaderLabels.MTDHRS)).getA1Notation()+')';
    let correctingFactorFormula = '=IFS('+condition1+','+then1+','+condition2+','+then2+')';
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.CORRECTFACTOR)).setFormula(correctingFactorFormula);
  }
}

function setStandardisedHours() {
  //correctingFactor * actualHours
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=2; i<=sheet.getLastRow(); i++) {
    let stdHrsFormula = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.ACTUALTIME)).getA1Notation()+'*'+sheet.getRange(i,HeaderIndex.get(HeaderLabels.CORRECTFACTOR)).getA1Notation()+')';
    sheet.getRange(i, HeaderIndex.get(HeaderLabels.STDTIME)).setFormula(stdHrsFormula);
  }
}

function setCost() {
  //dur*24*hourlyCost
  //setHeaderIndex();
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  for(let i=2; i<=sheet.getLastRow(); i++) {
    let costFormula = '('+sheet.getRange(i,HeaderIndex.get(HeaderLabels.STDTIME)).getA1Notation()+'*'+Configs.COSTFACTOR+')';
    sheet.getRange(i,HeaderIndex.get(HeaderLabels.COST)).setFormula(costFormula);
  }
}

function setDurationFormat() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  setHeaderIndex();
  /* for each duration column / just the used one, StandardisedHours */
  let range = sheet.getRange(2, HeaderIndex.get(HeaderLabels.STDTIME),sheet.getLastRow()-1);
  let durFormat = '[h]:mm:ss';
  range.setNumberFormat(durFormat)
}

function setEntity(brandIndex, entity) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  //setHeaderIndex();
  /*
   * Arrange by brand, find the range and apply the brand to the region.
   * 'entity' = brand: {'ETA', 'P529'}
   * 'brandIndex' = Headerinde.get(HeaderLabel.BRAND)
   */

  let filter = sheet.getFilter();
  if(filter) {
    filter.remove();
  }

  let start = -1;
  let end = -1;

  let range = sheet.getRange(1,1,sheet.getLastRow(), sheet.getLastColumn());
  filter = range.createFilter();
  filter.sort(brandIndex,true);
  for(let i=1; i<=sheet.getLastRow(); i++) {
    if(sheet.getRange(i,brandIndex).getValue() == entity) {
      start = i;
      break;
    }
  }
  for(let i=start; i<=sheet.getLastRow(); i++) {
    if(sheet.getRange(i,brandIndex).getValue() != entity) {
      end = i-1;
      break;
    }
  }

  for(start; start<=end; start++) {
    range = sheet.getRange(start, HeaderIndex.get(HeaderLabels.REGION));
    //Logger.log(range.getA1Notation());
    range.setValue(entity);
  }
  filter.remove();

}

function setBrandIndex() {
  /*
   * No for a set enum.
   * Brand to Region hardcoded in main().
   */
}

function main() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheet);
  /**
   * Pass bool isEndOfMonth to setConfigs, which then determines the dates set.
   * I.e is it the previous month, or current month.
   */
  const isEndOfMonth = false;

  setConfigs(isEndOfMonth);
  Logger.log('configs set.');
  setHeaderIndex();
  Logger.log('header index set.');
  clearZeroEst();
  Logger.log('zeros cleared.');
  clearLastModified();
  Logger.log('modified cleared.');
  clearCompletedAt();
  Logger.log('completed at cleared.');
  separateSharedTasks();
  Logger.log('separated tasks.');
  formatDev();
  Logger.log('formatted dev.');
  setActualTime();
  Logger.log('actual time set.');
  setSumOfActualTime();
  Logger.log('summed set.');
  setMonthToDateHours();
  Logger.log('mtd hrs set.');
  setCorrectingfactor();
  Logger.log('correcting factor set.');
  setStandardisedHours();
  Logger.log('standardised set.');
  setDurationFormat();
  Logger.log('dur set.');
  let brandIndex = HeaderIndex.get(HeaderLabels.BRAND);
  let entity = 'ETA';
  setEntity(brandIndex,entity);
  Logger.log('entity set.');
  entity = 'P529';
  setEntity(brandIndex,entity);
  Logger.log('entity set.');
  setCost();
  Logger.log('cost set.');
}

function aryna() {
  /**
   * Set header index
   * Clear completed before start of 24
   * Clear zero estimates
   * Split shared tasks
   * Format devs
   * Set brand index
   * Set brand entities
   */
  setHeaderIndex();
  Logger.log('index set');
  Configs.STARTDATE = new Date('January 01, 2024 01:00:00');
  Logger.log(Configs.STARTDATE);
  Configs.ENDDATE = new Date('December 31, 2024 01:00:00');
  Logger.log(Configs.ENDDATE);
  clearCompletedAt();
  Logger.log('completed at');
  clearZeroEst();
  Logger.log('zero est');
  separateSharedTasks();
  Logger.log('shared');
  formatDev();
  Logger.log('format');
  setBrandIndex(HeaderIndex.get(HeaderLabels.BRAND),'ETA');
  Logger.log('ETA');
  setBrandIndex(HeaderIndex.get(HeaderLabels.BRAND),'P529');
  Logger.log('P');
}