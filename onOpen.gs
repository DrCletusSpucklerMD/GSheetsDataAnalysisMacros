/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 * 
 * @currentDocOnly
 * 
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Special Functions')
    .addItem('Sort Results', 'sortResults')
    .addSeparator()
    .addItem('Filter Results', 'filterResults')
    .addSeparator()
    .addItem('Create Pivot Tables', 'pivot')
    .addSeparator()
    .addItem('Print Day 1 AS Bottles', 'printASDay1')
    .addSeparator()
    .addItem('Print Day 2 AS Bottles', 'printASDay2')
    .addSeparator()
    .addItem('Print Day 3 AS Bottles', 'printASDay3')
    .addSeparator()
    .addItem('Print SASEA Bottles', 'printSASEA')
    .addSeparator()
    .addItem('Print PL Bottles', 'printPL')
    .addSeparator()
    .addItem('Print Ladder', 'printLadder')
    .addSeparator()
    .addItem('Print Day 1 USD Bottles', 'printUSDDay1')
    .addSeparator()
    .addItem('Print Day 2 USD Bottles', 'printUSDDay2')
    .addSeparator()
    .addItem('Print Day 3 USD Bottles', 'printUSDDay3')
    .addSeparator()
    .addItem('Clear All', 'clearAll')
    .addToUi();
}
