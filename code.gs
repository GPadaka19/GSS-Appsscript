function copyAndPasteData() {
  // Buka spreadsheet aktif dan sheet "Master"
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Master');

  // Ambil data dari rentang J2:V2355
  var dataRange = sheet.getRange('J2:V2355');
  var data = dataRange.getValues();

  // Buat array baru untuk menyimpan data dalam format satu kolom tanpa nilai kosong
  var newData = [];
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] !== null && data[i][j] !== '') {
        newData.push([data[i][j]]);
      }
    }
  }

  // Tempelkan data ke kolom W secara menurun
  var targetRange = sheet.getRange(2, 23, newData.length, 1);
  targetRange.setValues(newData);

  // Ambil data dari kolom W untuk mendapatkan nilai unik
  var columnWRange = sheet.getRange(2, 23, newData.length, 1);
  var columnWData = columnWRange.getValues();

  // Buat array untuk menyimpan nilai unik
  var uniqueData = [];
  var seen = new Set();
  for (var k = 0; k < columnWData.length; k++) {
    var value = columnWData[k][0];
    if (value !== '' && !seen.has(value)) {
      seen.add(value);
      uniqueData.push([value]);
    }
  }

  // Tempelkan data unik ke kolom X secara menurun
  var uniqueTargetRange = sheet.getRange(2, 24, uniqueData.length, 1);
  uniqueTargetRange.setValues(uniqueData);

  // Ubah data dari kolom X menjadi uppercase dan tempelkan ke kolom Y
  var columnXRange = sheet.getRange(2, 24, uniqueData.length, 1);
  var columnXData = columnXRange.getValues();
  
  var uppercaseData = columnXData.map(function(row) {
    return [row[0].toUpperCase()];
  });

  var uppercaseTargetRange = sheet.getRange(2, 25, uppercaseData.length, 1);
  uppercaseTargetRange.setValues(uppercaseData);

  // Ambil data dari kolom Y untuk mendapatkan nilai unik
  var columnYRange = sheet.getRange(2, 25, uppercaseData.length, 1);
  var columnYData = columnYRange.getValues();

  // Buat array untuk menyimpan nilai unik
  var uniqueYData = [];
  var seenY = new Set();
  for (var m = 0; m < columnYData.length; m++) {
    var valueY = columnYData[m][0];
    if (valueY !== '' && !seenY.has(valueY)) {
      seenY.add(valueY);
      uniqueYData.push([valueY]);
    }
  }

  // Tempelkan data unik ke kolom Z secara menurun
  var uniqueYTargetRange = sheet.getRange(2, 26, uniqueYData.length, 1);
  uniqueYTargetRange.setValues(uniqueYData);
}
