function createAndRenameSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = spreadsheet.getSheetByName("Master"); // Ubah sesuai dengan nama sheet master Anda
  
  // Array nama sheet dan nilai filter
  var sheetsData = [
    { name: "Room A", filterValue: "Room A", headerColumns: ["A1:H1", "R1:S1"]  },
    { name: "Room B", filterValue: "Room B", headerColumns: ["A1:G1", "I1", "R1:S1"]  },
    { name: "Room C", filterValue: "Room C", headerColumns: ["A1:G1", "J1", "R1:S1"]  },
    { name: "Room D", filterValue: "Room D", headerColumns: ["A1:G1", "K1", "R1:S1"]  },
    { name: "Room E", filterValue: "Room E", headerColumns: ["A1:G1", "L1", "R1:S1"]  },
    { name: "Room F", filterValue: "Room F", headerColumns: ["A1:G1", "M1", "R1:S1"]  },
    { name: "Room G", filterValue: "Room G", headerColumns: ["A1:G1", "N1", "R1:S1"]  },
    { name: "Room H", filterValue: "Room H", headerColumns: ["A1:G1", "O1", "R1:S1"]  },
    { name: "Room I", filterValue: "Room I", headerColumns: ["A1:G1", "P1", "R1:S1"]  },
    { name: "Room J", filterValue: "Room J", headerColumns: ["A1:G1", "Q1:S1"]  },
    { name: "Room K", filterValue: "Room K", headerColumns: ["A1:G1", "R1", "R1:S1"]  },
    { name: "Room L", filterValue: "Room L", headerColumns: ["A1:G1", "S1", "R1:S1"]  },
    { name: "Room M", filterValue: "Room M", headerColumns: ["A1:G1", "T1", "R1:S1"]  },
    { name: "Room N", filterValue: "Room N", headerColumns: ["A1:G1", "U1", "R1:S1"]  },
    { name: "Room O", filterValue: "Room O", headerColumns: ["A1:G1", "V1", "R1:S1"]  }
  ];

  // Loop untuk membuat dan menamai sheet baru
  for (var i = 0; i < sheetsData.length; i++) {
    var sheetName = sheetsData[i].name;
    var filterValue = sheetsData[i].filterValue;
    var headerColumns = sheetsData[i].headerColumns;
    
    var newSheet = spreadsheet.insertSheet(sheetName); // Membuat sheet baru
    var range = newSheet.getRange("A2"); // Mendapatkan range cell A2 di sheet baru
    range.setFormula('=FILTER(Master!A:V, Master!G:G="' + filterValue + '")'); // Set rumus pada cell A2
    
    // Menyalin dan menempelkan data dari sel header di sheet master ke setiap sheet baru
    for (var j = 0; j < headerColumns.length; j++) {
      var sourceRange = masterSheet.getRange(headerColumns[j]);
      var targetRange = newSheet.getRange(headerColumns[j]);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES);
    }

    // Menyembunyikan kolom yang memiliki header kosong
    var headerRange = newSheet.getRange("A1:V1"); // Anggap header berada di kolom A hingga V
    var headers = headerRange.getValues()[0]; // Dapatkan nilai header
    for (var k = 0; k < headers.length; k++) {
      if (headers[k] == "") { // Jika header kosong
        newSheet.hideColumns(k + 1); // Sembunyikan kolom (k+1 karena indeks dimulai dari 0)
      }
    }
  }
}
