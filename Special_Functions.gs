/**
 * TODO:
 *  add gradients to pivot table
 *
 * unlableed?
 *
 */

//Defines Spreadsheet
let plateMapSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

//Defines Sheets
const plateMap96Sheet = plateMapSpreadsheet.getSheetByName("96-Well Plates");
const plateMap384Sheet = plateMapSpreadsheet.getSheetByName("384 Layout");
const regressionSheet = plateMapSpreadsheet.getSheetByName("regression");
const singleThresholdSheet = plateMapSpreadsheet.getSheetByName("single_threshold");
const positivesSheet = plateMapSpreadsheet.getSheetByName("positives");
const resultsListSheet = plateMapSpreadsheet.getSheetByName("RTL");
const saseaSheet = plateMapSpreadsheet.getSheetByName("SASEA");
const plSheet = plateMapSpreadsheet.getSheetByName("PL");
const sfoSheet = plateMapSpreadsheet.getSheetByName("SFO");

//Defines Ranges
const plateMap384Range = plateMap384Sheet.getRange("B2:Y17");
const regressionCsvRange = regressionSheet.getRange("E2:T1537");
const regressionCsvRangeWithHeaders = regressionSheet.getRange("A1:M1537");
const singleThresholdCsvRange = singleThresholdSheet.getRange("E2:T1537");
const singconsthresholdCsvRangeWithHeaders = regressionSheet.getRange("E1:T1537");
const asResultsRangeDay1 = resultsListSheet.getRange("B1:B144");
const asResultsRangeDay2 = resultsListSheet.getRange("E1:E144");
const asResultsRangeDay3 = resultsListSheet.getRange("H1:H144");
const saseaResultsRange = saseaSheet.getRange("A2:AI2");
const plResultsRange = plSheet.getRange("A2:X2");
const sfoResultsRange = sfoSheet.getRange("A2:V400");
const cqRange = positivesSheet.getRange("D:F");
const rawStData = positivesSheet.getRange('single_threshold!A1:T1537');
const rawRegData = positivesSheet.getRange('regression!A1:T1537');


//Defines Range Values
const ogPlateMapList = plateMap384Range.getValues();
let modPlateMapList = [];
ogPlateMapList.forEach(entry => {
  let tempArray = entry.toString().split(",");
  tempArray.forEach(subentry => {
    if (subentry != "") {
      modPlateMapList.push(subentry);
    }
  })
})
uniq = [...new Set(modPlateMapList)];

const resultsListDay1 = asResultsRangeDay1.getValues();
const resultsListDay2 = asResultsRangeDay2.getValues();
const resultsListDay3 = asResultsRangeDay3.getValues();
const saseaList = saseaResultsRange.getValues().toString().split(",");

//Defines Filter Criteria
const fluoroFilter = SpreadsheetApp.newFilterCriteria().setHiddenValues(["Texas Red", "Cy5"]);
const cqFilter = SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(0);


/**
 * clearSheet
 *
 * method clears the sheets of previously inputted data
 *
 */

function clearAll(){

  regressionCsvRange.offset(0,1).clear();
  if (regressionSheet.getFilter() != null) {
    regressionSheet.getFilter().remove();
  }
  singleThresholdCsvRange.offset(0,1).clear();
  if (singleThresholdSheet.getFilter() != null) {
    singleThresholdSheet.getFilter().remove();
  }
  asResultsRangeDay1.offset(0,1).clear();
  asResultsRangeDay2.offset(0,1).clear();
  asResultsRangeDay3.offset(0,1).clear();
  saseaResultsRange.offset(1,0).clear();
  plResultsRange.clear();
  sfoResultsRange.clear();
  positivesSheet.getRange("D3:H3").clear();
  plateMap96Sheet.getRange("B4:M11").clearContent();
  plateMap96Sheet.getRange("B15:M22").clearContent();
  plateMap96Sheet.getRange("B26:M33").clearContent();
  plateMap96Sheet.getRange("B1:D1").setValues([["NA","NA","NA"]]);

}

/**
 * sortResults
 *
 * Function which sorts regression and single threshold CSV range
 *
 */

function sortResults() {
  singleThresholdCsvRange.sort([{column: 12, ascending: false}]);
  regressionCsvRange.sort([{column: 12, ascending: false}]);

}


/**
 * filterResults
 *
 * filters regression and single threshold results
 *
 */

function filterResults() {

  if (!regressionSheet.getRange("F1:T1537").getFilter()) {
  const regressionFilter = regressionSheet.getRange("F1:T1537").createFilter();
  regressionFilter.setColumnFilterCriteria(7,fluoroFilter);
  regressionFilter.setColumnFilterCriteria(13, cqFilter);
  } else {
    regressionSheet.getRange("F1:T1537").getFilter().remove();
  }

if (!singleThresholdSheet.getRange("F1:T1537").getFilter()) {
  const singleThresholdFilter = singleThresholdSheet.getRange("F1:T1537").createFilter();
  singleThresholdFilter.setColumnFilterCriteria(7,fluoroFilter);
  singleThresholdFilter.setColumnFilterCriteria(13, cqFilter);
} else {
  singleThresholdSheet.getRange("F1:T1537").getFilter().remove();
}

}

/**
 * pivot - creates pivot tables
 *
 * creates summary tables from regression and single threshold qPCR data
 *
 */

function pivot() {


  //builds pivot table based on regression analysis data
  const regPivotTable = positivesSheet.getRange('D3').createPivotTable(rawRegData);
  regPivotTable.addRowGroup(3).showTotals(false);
  const onlyFamHexCriteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['FAM', 'HEX'])
  .build();
  regPivotTable.addFilter(7, onlyFamHexCriteria);
  pivotValue = regPivotTable.addPivotValue(13, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  const greaterThanZeroCriteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberGreaterThan(0)
  .build()
  regPivotTable.addFilter(12, greaterThanZeroCriteria);


  //builds pivot table based off single threshold analysis data
  const stPivotTable = positivesSheet.getRange('G3').createPivotTable(rawStData);
  stPivotTable.addRowGroup(3).showTotals(false);
  stPivotTable.addFilter(7, onlyFamHexCriteria);
  pivotValue = stPivotTable.addPivotValue(13, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  stPivotTable.addFilter(12, greaterThanZeroCriteria);

}


/**
 * printASDay1
 *
 * finds AS bottles and Cq results for day 1
 *
 */

function printASDay1() {

  i = 0;
  console.log(resultsListDay1);
  resultsListDay1.forEach(result => {
    if (plateMap96Sheet.getRange("B:M").createTextFinder(result).findNext()) {
      if (positivesSheet.getRange("D:F").createTextFinder(result).findNext()) {
        let cq = cqRange.createTextFinder(result).findNext().offset(0,1).getValue();
        asResultsRangeDay1.offset(i,1).setValue(cq);
      } else {
        asResultsRangeDay1.offset(i,1).setValue("-1");
      }
    } else {
      asResultsRangeDay1.offset(i,1).setValue("");
    }
    i++;
  })
  asResultsRangeDay1.offset(i,1).setValue("");

}

/**
 * printASDay2
 *
 * finds AS bottles and Cq results for day 2
 *
 */

function printASDay2() {
  j = 0;
  resultsListDay2.forEach(result => {
    if (plateMap96Sheet.getRange("B:M").createTextFinder(result).findNext()) {
      if (positivesSheet.getRange("D:F").createTextFinder(result).findNext()) {
        let cq = cqRange.createTextFinder(result).findNext().offset(0,1).getValue();
        asResultsRangeDay2.offset(j,1).setValue(cq);
      } else {
        asResultsRangeDay2.offset(j,1).setValue("-1");
      }
    } else {
      asResultsRangeDay2.offset(j,1).setValue("");
    }
    j++;
  })
  asResultsRangeDay2.offset(j,1).setValue("");

}

/**
 * printASDay3
 *
 * finds AS bottles and Cq results for day 3
 *
 */

function printASDay3() {
  k = 0;
  resultsListDay3.forEach(result => {
    if (plateMap96Sheet.getRange("B:M").createTextFinder(result).findNext()) {
      if (positivesSheet.getRange("D:F").createTextFinder(result).findNext()) {
        let cq = cqRange.createTextFinder(result).findNext().offset(0,1).getValue();
        asResultsRangeDay3.offset(k,1).setValue(cq);
      } else {
        asResultsRangeDay3.offset(k,1).setValue("-1");
      }
    } else {
      asResultsRangeDay3.offset(k,1).setValue("");
    }
    k++;
  })
  asResultsRangeDay3.offset(k,1).setValue("");

}


/**
 * printSASEA
 *
 * uses REGEX to match school code to retrieve regression Cq results
 *
 */

function printSASEA() {

  l = 0;

  saseaList.forEach(result => {
    let reResult = String(result + "[^a-z]");
    if (plateMap96Sheet.getRange("B:M").createTextFinder(reResult).useRegularExpression(true).findNext() != null) {
      if (positivesSheet.getRange("D:E").createTextFinder(reResult).useRegularExpression(true).findNext() != null) {
        let cq = positivesSheet.getRange("D:E").createTextFinder(reResult).useRegularExpression(true).findNext().offset(0,1).getValue();
        saseaResultsRange.offset(1,l).setValue(cq);
        l++;
      } else {
        saseaResultsRange.offset(1, l).setValue("0");
        l++;
      }
    } else {
      saseaResultsRange.offset(1, l).setValue("NS");
      l++;
    }
  })
  saseaResultsRange.offset(1,l).setValue("");
}

/**
 * printPL
 *
 * creates a pivot table that organizes the Point Loma data into an easy to read manner
 *
 */

function printPL(){

  //builds pivot tables for PL data

  //st table
  const plStPivot = plSheet.getRange('A2').createPivotTable(rawStData);
  plStPivot.addColumnGroup(5).showTotals(false);
  onlyPromegaCriteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['*Promega - N1', '*Promega - N2', 'Promega - E'])
  .build();
  plStPivot.addFilter(5, onlyPromegaCriteria);
  const noTexCriteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['FAM','HEX','Cy5'])
  .build();
  plStPivot.addFilter(7, noTexCriteria);
  plStPivot.addColumnGroup(4).sortDescending().showTotals(false);
  plStPivot.addPivotValue(13, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  plStPivot.addRowGroup(3).showTotals(false);
  const noBlanksCriteria = SpreadsheetApp.newFilterCriteria()
  .whenCellNotEmpty()
  .build();
  plStPivot.addFilter(3, noBlanksCriteria);

  //reg table
  const plRegPivot = plSheet.getRange('K2').createPivotTable(rawRegData);
  plRegPivot.addColumnGroup(5).showTotals(false);
  onlyPromegaCriteria = SpreadsheetApp.newFilterCriteria()
  .setVisibleValues(['*Promega - N1', '*Promega - N2', 'Promega - E'])
  .build();
  plRegPivot.addFilter(5, onlyPromegaCriteria);
  plRegPivot.addFilter(7, noTexCriteria);
  plRegPivot.addColumnGroup(4).sortDescending().showTotals(false);
  plRegPivot.addPivotValue(13, SpreadsheetApp.PivotTableSummarizeFunction.AVERAGE);
  plRegPivot.addRowGroup(3).showTotals(false);
  plRegPivot.addFilter(3, noBlanksCriteria);

}
