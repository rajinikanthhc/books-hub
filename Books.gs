// ==========================================
// BOOK FUNCTIONS
// ==========================================


/* ==========================================
   GET BOOKS
========================================== */

function getBooks() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Sheet not found: " + CONFIG.SHEET_NAME
    );
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  values.shift();

  return values

    .filter(row => row[1])

    .map(row => ({

      id: row[0],

      title: row[1],

      author: row[2],

      language: row[3],

      category: row[4],

      pdf: row[5],

      cover: getCoverUrl(row[6])

    }));

}


/* ==========================================
   COVER URL
========================================== */

function getCoverUrl(value) {

  if (!value) {

    return (
      CONFIG.COVER_BASE_URL +
      encodeURIComponent(
        CONFIG.DEFAULT_COVER
      )
    );

  }

  if (
    typeof value === "string" &&
    value.startsWith("http")
  ) {

    return value;

  }

  return (
    CONFIG.COVER_BASE_URL +
    encodeURIComponent(value)
  );

}


/* ==========================================
   LOWEST AVAILABLE ID
   STARTS FROM 1
========================================== */

function getNextBookId() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_NAME);

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return 1;
  }

  const ids =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat()
      .map(id => Number(id))
      .filter(id =>
        Number.isInteger(id) &&
        id >= 1
      );

  let id = 1;

  while (ids.includes(id)) {
    id++;
  }

  return id;

}


/* ==========================================
   ADD BOOK
========================================== */

function addBook(bookData) {

  if (!bookData.title) {

    throw new Error(
      "Book title is required."
    );

  }

  if (!bookData.pdfData) {

    throw new Error(
      "Please select the PDF."
    );

  }

  if (!bookData.coverData) {

    throw new Error(
      "Please select the cover image."
    );

  }


  /* ----------------------------------------
     LOWEST AVAILABLE ID
  ---------------------------------------- */

  const id =
    getNextBookId();


  /* ----------------------------------------
     UPLOAD PDF TO GOOGLE DRIVE
  ---------------------------------------- */

  const pdfBytes =
    Utilities.base64Decode(
      bookData.pdfData
    );

  const pdfBlob =
    Utilities.newBlob(
      pdfBytes,
      bookData.pdfMimeType ||
        "application/pdf",
      bookData.pdfName
    );


  const booksFolder =
    DriveApp.getFolderById(
      CONFIG.BOOKS_FOLDER_ID
    );


  const pdfFile =
    booksFolder.createFile(
      pdfBlob
    );


  pdfFile.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );


  const pdfUrl =
    "https://drive.google.com/file/d/" +
    pdfFile.getId() +
    "/view";


  /* ----------------------------------------
     AUTOMATIC COVER NAME
  ---------------------------------------- */

  const coverName =
    createCoverFileName(
      bookData.title,
      bookData.coverName,
      bookData.coverMimeType
    );


  /* ----------------------------------------
     UPLOAD COVER TO GITHUB
  ---------------------------------------- */

  const coverUrl =
    uploadCoverToGitHub(
      coverName,
      bookData.coverData
    );


  /* ----------------------------------------
     SAVE TO SHEET
  ---------------------------------------- */

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEET_NAME
      );


  sheet.appendRow([

    id,

    bookData.title,

    bookData.author || "",

    bookData.language || "English",

    bookData.category || "",

    pdfUrl,

    coverName,

    ""

  ]);


  return {

    success: true,

    id: id,

    cover: coverUrl

  };

}


/* ==========================================
   CREATE COVER FILE NAME
   BOOK TITLE + ORIGINAL EXTENSION
========================================== */

function createCoverFileName(
  title,
  originalName,
  mimeType
) {

  /* ----------------------------------------
     GET EXTENSION
  ---------------------------------------- */

  let extension = "";


  if (originalName) {

    const match =
      originalName.match(
        /\.([a-zA-Z0-9]+)$/
      );

    if (match) {
      extension =
        "." +
        match[1].toLowerCase();
    }

  }


  /* ----------------------------------------
     FALLBACK FROM MIME TYPE
  ---------------------------------------- */

  if (!extension && mimeType) {

    const mimeMap = {

      "image/jpeg": ".jpg",

      "image/jpg": ".jpg",

      "image/png": ".png",

      "image/webp": ".webp",

      "image/gif": ".gif",

      "image/bmp": ".bmp",

      "image/svg+xml": ".svg"

    };

    extension =
      mimeMap[mimeType] || ".jpg";

  }


  /* ----------------------------------------
     CLEAN BOOK TITLE
  ---------------------------------------- */

  let cleanTitle =
    String(title)
      .trim()
      .replace(
        /[<>:"/\\|?*]+/g,
        ""
      )
      .replace(
        /\s+/g,
        " "
      );


  /* ----------------------------------------
     FINAL FILE NAME
  ---------------------------------------- */

  return (
    cleanTitle +
    extension
  );

}


/* ==========================================
   GITHUB COVER UPLOAD
========================================== */

function uploadCoverToGitHub(
  fileName,
  base64Data
) {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "GITHUB_TOKEN"
      );


  if (!token) {

    throw new Error(
      "GITHUB_TOKEN is not configured."
    );

  }


  const url =
    "https://api.github.com/repos/" +
    CONFIG.GITHUB_OWNER +
    "/" +
    CONFIG.GITHUB_REPO +
    "/contents/" +
    CONFIG.GITHUB_COVER_FOLDER +
    "/" +
    encodeURIComponent(
      fileName
    );


  /* ----------------------------------------
     CHECK IF FILE ALREADY EXISTS
  ---------------------------------------- */

  const existingResponse =
    UrlFetchApp.fetch(
      url,
      {

        method: "get",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        muteHttpExceptions: true

      }
    );


  const existingCode =
    existingResponse.getResponseCode();


  let sha = null;


  if (existingCode === 200) {

    const existingFile =
      JSON.parse(
        existingResponse.getContentText()
      );

    sha =
      existingFile.sha;

  }


  /* ----------------------------------------
     PAYLOAD
  ---------------------------------------- */

  const payload = {

    message:
      "Add book cover: " +
      fileName,

    content:
      base64Data

  };


  if (sha) {

    payload.sha = sha;

  }


  /* ----------------------------------------
     UPLOAD
  ---------------------------------------- */

  const response =
    UrlFetchApp.fetch(
      url,
      {

        method: "put",

        contentType:
          "application/json",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions: true

      }
    );


  const code =
    response.getResponseCode();


  const text =
    response.getContentText();


  if (
    code !== 200 &&
    code !== 201
  ) {

    throw new Error(
      "GitHub upload failed.\n\n" +
      "Response: " +
      text
    );

  }


  return (
    "https://raw.githubusercontent.com/" +
    CONFIG.GITHUB_OWNER +
    "/" +
    CONFIG.GITHUB_REPO +
    "/main/" +
    CONFIG.GITHUB_COVER_FOLDER +
    "/" +
    encodeURIComponent(
      fileName
    )
  );

}


/* ==========================================
   DELETE BOOK
   PASSCODE REQUIRED
========================================== */

function deleteBook(
  id,
  passcode
) {

  /* ----------------------------------------
     DELETE PASSCODE
  ---------------------------------------- */

  if (
    String(passcode) !==
    "12345"
  ) {

    throw new Error(
      "Incorrect delete passcode."
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEET_NAME
      );


  const data =
    sheet
      .getDataRange()
      .getValues();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      Number(data[i][0]) ===
      Number(id)
    ) {


      /* --------------------------------
         DELETE PDF FROM DRIVE
      -------------------------------- */

      try {

        const pdfUrl =
          data[i][5];

        const pdfId =
          extractDriveFileId(
            pdfUrl
          );


        if (pdfId) {

          DriveApp
            .getFileById(pdfId)
            .setTrashed(true);

        }

      } catch (e) {

        console.log(
          "PDF delete error: " +
          e
        );

      }


      /* --------------------------------
         DELETE COVER FROM GITHUB
      -------------------------------- */

      try {

        const coverName =
          data[i][6];


        if (coverName) {

          deleteCoverFromGitHub(
            coverName
          );

        }

      } catch (e) {

        console.log(
          "Cover delete error: " +
          e
        );

      }


      /* --------------------------------
         DELETE SHEET ROW
      -------------------------------- */

      sheet.deleteRow(
        i + 1
      );


      return {

        success: true

      };

    }

  }


  throw new Error(
    "Book not found."
  );

}


/* ==========================================
   DELETE COVER FROM GITHUB
========================================== */

function deleteCoverFromGitHub(
  fileName
) {

  const token =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        "GITHUB_TOKEN"
      );


  if (!token) {

    throw new Error(
      "GITHUB_TOKEN is not configured."
    );

  }


  const url =
    "https://api.github.com/repos/" +
    CONFIG.GITHUB_OWNER +
    "/" +
    CONFIG.GITHUB_REPO +
    "/contents/" +
    CONFIG.GITHUB_COVER_FOLDER +
    "/" +
    encodeURIComponent(
      fileName
    );


  /* --------------------------------
     GET FILE SHA
  -------------------------------- */

  const getResponse =
    UrlFetchApp.fetch(
      url,
      {

        method: "get",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        muteHttpExceptions: true

      }
    );


  const getCode =
    getResponse.getResponseCode();


  if (getCode !== 200) {

    return;

  }


  const fileInfo =
    JSON.parse(
      getResponse.getContentText()
    );


  /* --------------------------------
     DELETE FILE
  -------------------------------- */

  const payload = {

    message:
      "Delete book cover: " +
      fileName,

    sha:
      fileInfo.sha

  };


  const response =
    UrlFetchApp.fetch(
      url,
      {

        method: "delete",

        contentType:
          "application/json",

        headers: {

          Authorization:
            "Bearer " + token,

          Accept:
            "application/vnd.github+json",

          "X-GitHub-Api-Version":
            "2022-11-28"

        },

        payload:
          JSON.stringify(
            payload
          ),

        muteHttpExceptions: true

      }
    );


  console.log(
    "GitHub delete response: " +
    response.getContentText()
  );

}


/* ==========================================
   DRIVE FILE ID
========================================== */

function extractDriveFileId(
  url
) {

  if (!url) {

    return null;

  }


  let match =
    url.match(
      /\/d\/([a-zA-Z0-9_-]+)/
    );


  if (match) {

    return match[1];

  }


  match =
    url.match(
      /[?&]id=([a-zA-Z0-9_-]+)/
    );


  if (match) {

    return match[1];

  }


  return null;

}