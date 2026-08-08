//==========================================
// BOOK FUNCTIONS
//==========================================

function getBooks() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {

    throw new Error("Books sheet not found.");

  }

  const values = sheet
    .getDataRange()
    .getValues();

  values.shift();

  return values.map(row => ({

    id: row[0],

    title: row[1],

    author: row[2],

    language: row[3],

    category: row[4],

    pdf: row[5],

    cover:
      CONFIG.COVER_BASE_URL +
      encodeURIComponent(
        row[6] || CONFIG.DEFAULT_COVER
      ),

    lastPage: row[7] || 0

  }));

}