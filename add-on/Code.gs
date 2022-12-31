let ui = SpreadsheetApp.getUi();
let sheet = SpreadsheetApp.getActiveSheet();
let row;
let status = "No Analysis Running";

function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  ui.createAddonMenu()
    .addItem("Start", "showSidebar")
    .addSeparator()
    .addItem("Help", "showHelp")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile("index").evaluate();
  html.setTitle("Data Analytics");
  ui.showSidebar(html);
}

function showHelp() {
  Logger.log("Viewing help");
  console.log("Log");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getColumnHeader() {
  let data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  const header = data[0].filter(Boolean);
  return header;
}

function getByName(columnName) {
  let columnValue = [];
  let data = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  let columnIndex = data[0].indexOf(columnName);

  for (row = 1; row < data.length; ++row) {
    columnValue.push(data[row][columnIndex]);
  }
  return columnValue;
}

function writeHeader(header, lastCol) {
  sheet
    .getRange(1, lastCol + 1, 1, header.length)
    .setValues([header])
    .setBackground("#9b4dca")
    .setFontColor("#ffffff");
}

function lastColumnValue() {
  let last = SpreadsheetApp.getActiveSheet().getLastColumn();
  return last;
}

function writeDataToRange(data, column, counter, lastCol) {
  row = counter + 1;
  sheet
    .getRange(row, lastCol + 1, data.length, column)
    .setValues(data)
    .setBackground("#f4f5f6");
  return status;
}
