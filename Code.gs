//==========================================
// BOOKS HUB
//==========================================

function doGet() {

  return HtmlService
    .createTemplateFromFile("Index")
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}


function include(fileName){

  return HtmlService
    .createHtmlOutputFromFile(fileName)
    .getContent();

}

function authorizeBooksHub() {

  const folder =
    DriveApp.getFolderById(
      CONFIG.BOOKS_FOLDER_ID
    );

  Logger.log(
    "Books folder authorized: " +
    folder.getName()
  );

}