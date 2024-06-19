function createAndRenameSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master"); // Ubah sesuai dengan nama sheet master Anda

  // Array nama sheet dan nilai filter
  var sheetsData = [
    { name: "Room Website A", filterValue: "Room Website A", headerColumns: ["A1:J1"] },
    { name: "Room Website B", filterValue: "Room Website B", headerColumns: ["A1:I1", "K1"] },
    { name: "Room Website C", filterValue: "Room Website C", headerColumns: ["A1:I1", "L1"] },
    { name: "Room Instagram A", filterValue: "Room Instagram A", headerColumns: ["A1:I1", "M1"] },
    { name: "Room Instagram B", filterValue: "Room Instagram B", headerColumns: ["A1:I1", "N1"] },
    { name: "Room Tiktok A", filterValue: "Room Tiktok A", headerColumns: ["A1:I1", "O1"] },
    { name: "Room Tiktok B", filterValue: "Room Tiktok B", headerColumns: ["A1:I1", "P1"] },
  ];

  // Loop untuk membuat dan menamai sheet baru
  for (var i = 0; i < sheetsData.length; i++) {
    var sheetName = sheetsData[i].name;
    var filterValue = sheetsData[i].filterValue;
    var headerColumns = sheetsData[i].headerColumns;

    var newSheet = spreadsheet.insertSheet(sheetName); // Membuat sheet baru
    var range = newSheet.getRange("A2"); // Mendapatkan range cell A2 di sheet baru
    range.setFormula('=FILTER(Master!A:P;Master!I:I="' + filterValue + '")'); // Set rumus pada cell A2

    // Menyalin dan menempelkan data dari sel header di sheet master ke setiap sheet baru
    for (var j = 0; j < headerColumns.length; j++) {
      var sourceRange = masterSheet.getRange(headerColumns[j]);
      var targetRange = newSheet.getRange(headerColumns[j]);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES);
    }

    // Menyembunyikan kolom yang memiliki header kosong
    var headerRange = newSheet.getRange("A1:P1"); // Anggap header berada di kolom A hingga P
    var headers = headerRange.getValues()[0]; // Dapatkan nilai header
    for (var k = 0; k < headers.length; k++) {
      if (headers[k] == "") { // Jika header kosong
        newSheet.hideColumns(k + 1); // Sembunyikan kolom (k+1 karena indeks dimulai dari 0)
      }
    }
  }
}
