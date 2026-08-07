// =====================================
// BOOKS HUB
// Backend
// =====================================

const SHEET_NAME = "Books Hub";

// Home Page
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Books Hub")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include HTML Files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Read Books
function getBooks() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  data.shift(); // Remove Header

  return data.map(row => ({

    id: row[0],
    title: row[1],
    author: row[2],
    language: row[3],
    category: row[4],
    pdf: row[5],
    lastPage: row[6]

  }));

}