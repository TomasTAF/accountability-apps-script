function setFontForAllSheets() {
  var font = 'Calibri'; // Cambia esto al tipo de fuente que prefieras
  var fontSize = 11; // Cambia esto al tamaño de letra que prefieras
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  sheets.forEach(function(sheet) {
    var range = sheet.getDataRange();
    range.setFontFamily(font);
    range.setFontSize(fontSize);
  });
}

function sumDebitCreditInSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNamePattern = /^\d{4}-\d{2}-\d{2} P[IED] \d$|^\d{4}-\d{2}-\d{2} P[IED] \d\d$/;

  sheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetNamePattern.test(sheetName)) {
      var lastRow = sheet.getLastRow();
      var rangeDebit = sheet.getRange(2, 5, lastRow - 2, 1); // Columna E desde la fila 2 hasta la última fila
      var rangeCredit = sheet.getRange(2, 6, lastRow - 2, 1); // Columna F desde la fila 2 hasta la última fila

      var sumDebit = rangeDebit.getValues().reduce(function(sum, value) {
        return sum + (parseFloat(value[0]) || 0);
      }, 0);

      var sumCredit = rangeCredit.getValues().reduce(function(sum, value) {
        return sum + (parseFloat(value[0]) || 0);
      }, 0);

      sheet.getRange(lastRow, 5).setValue(sumDebit); // Suma en la columna E
      sheet.getRange(lastRow, 6).setValue(sumCredit); // Suma en la columna F
    }
  });
}

function updateLedgerWithPolicy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ledgerSheet = ss.getSheetByName('Libro Mayor'); // Asegúrate de que el nombre de la hoja sea correcto
  var condensedSheet = ss.getSheetByName('Condensado'); // Asegúrate de que el nombre de la hoja sea correcto

  if (ledgerSheet && condensedSheet) {
    var lastRowLedger = ledgerSheet.getLastRow();
    var lastRowCondensed = condensedSheet.getLastRow();
    var policyRange = condensedSheet.getRange(2, 2, lastRowCondensed - 1, 1); // Columna B desde la fila 2 hasta la última fila

    var policies = policyRange.getValues();

    for (var i = 2; i <= lastRowLedger; i++) {
      var policy = policies[i - 2][0]; // Ajusta el índice según sea necesario
      ledgerSheet.getRange(i, 3).setValue(policy); // Columna C en el libro mayor
    }
  }
}
