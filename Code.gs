const SHEET_NAME = "Responses";

const HEADERS = [
  "Timestamp",
  "Nama",
  "Kelas",
  "Nama Feedback Dipilih",
  "Versi 1 - Isi",
  "Versi 1 - Bahasa & Organisasi",
  "Versi 1 - Perkara Baik 1",
  "Versi 1 - Perkara Baik 2",
  "Versi 1 - Perlu Baiki 1",
  "Versi 1 - Perlu Baiki 2",
  "Versi 2 - Isi",
  "Versi 2 - Bahasa & Organisasi",
  "Perkara Baiki 1 Sudah Dibetulkan?",
  "Perkara Baiki 2 Sudah Dibetulkan?",
  "Versi 2 - Perkara Baik",
  "Versi 2 - Perkara Lain Perlu Dibaiki",
  "Renungan 1",
  "Renungan 2"
];

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#6f2f3c")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  const p = e.parameter || {};

  const row = [
    new Date(),
    p.studentName || "",
    p.studentClass || "",
    p.feedbackStudent || "",
    p.oldContent || "",
    p.oldLanguage || "",
    p.goodThing1 || "",
    p.goodThing2 || "",
    p.improve1 || "",
    p.improve2 || "",
    p.newContent || "",
    p.newLanguage || "",
    p.improve1Fixed || "",
    p.improve2Fixed || "",
    p.newGood1 || "",
    p.newImprove1 || "",
    p.reflection1 || "",
    p.reflection2 || ""
  ];

  sheet.appendRow(row);
  sheet.autoResizeColumns(1, HEADERS.length);

  return ContentService
    .createTextOutput(JSON.stringify({status: "success"}))
    .setMimeType(ContentService.MimeType.JSON);
}
