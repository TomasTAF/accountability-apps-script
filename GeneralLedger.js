function processAccount(accountCodePrefix, mayorSheet, startRow) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Consolidado');

  if (!sheet) {
    Logger.log('La hoja "Consolidado" no se encontró.');
    return startRow;
  }

  if (!mayorSheet) {
    Logger.log('La hoja mayor no se encontró.');
    return startRow;
  }

  var data = sheet.getDataRange().getValues();
  var totalDebe = 0;
  var totalHaber = 0;

  // Escribir la descripción de la cuenta y ponerla en negritas
  var accountRange = mayorSheet.getRange(startRow, 1, 1, 1);
  accountRange.setValue('Cuenta ' + accountCodePrefix);
  accountRange.setFontWeight('bold');
  startRow++;

  var firstDebitRow = -1;

  for (var i = 1; i < data.length; i++) {
    var accountCode = data[i][2];
    if (accountCode && accountCode.startsWith(accountCodePrefix)) {
      var debito = parseFloat(data[i][5]) || 0;
      var credito = parseFloat(data[i][6]) || 0;
      var poliza = data[i][1]; // Extraer el valor de la segunda columna como póliza

      if (debito === 0 && credito === 0) {
        continue;
      }

      totalDebe += debito;
      totalHaber += credito;

      var rowRange = mayorSheet.getRange(startRow, 1, 1, 3);
      rowRange.setValues([[debito, credito, poliza]]); // Colocar la póliza en la tercera columna
      
      if (debito !== 0 && firstDebitRow === -1) {
        firstDebitRow = startRow;
      } else {
        // Aplicar borde derecho a las celdas posteriores del debe
        if (debito !== 0) {
          var debeCell = mayorSheet.getRange(startRow, 1);
          debeCell.setBorder(false, false, false, true, false, false);
        }
        // Aplicar borde izquierdo a las celdas posteriores del haber
        if (credito !== 0) {
          var haberCell = mayorSheet.getRange(startRow, 2);
          haberCell.setBorder(false, true, false, false, false, false);
        }
      }

      startRow++;
    }
  }

  // Añadir fila de totalización con bordes
  var totalRange = mayorSheet.getRange(startRow, 1, 1, 3);
  totalRange.setValues([[totalDebe, totalHaber, ""]]);
  totalRange.getCell(1, 1).setBorder(true, false, false, true, false, false); // Borde superior y derecho en el total del debito
  totalRange.getCell(1, 2).setBorder(true, true, false, false, false, false); // Borde superior e izquierdo en el total del credito
  startRow++;

  var diferencia = totalDebe - totalHaber;
  if (totalDebe >= totalHaber) {
    var diferenciaCell = mayorSheet.getRange(startRow, 1);
    diferenciaCell.setValue(diferencia);
    diferenciaCell.setBorder(true, false, false, true, false, false); // Borde superior y derecho en el saldo del debe
    mayorSheet.getRange(startRow, 2).setValue('');
  } else {
    var diferenciaCell = mayorSheet.getRange(startRow, 2);
    diferenciaCell.setValue(diferencia);
    diferenciaCell.setBorder(true, true, false, false, false, false); // Borde superior e izquierdo en el saldo del haber
    mayorSheet.getRange(startRow, 1).setValue('');
  }
  startRow++;

  // Añadir dos filas en blanco
  startRow += 2;

  // Aplicar bordes a las primeras celdas de debe y haber
  if (firstDebitRow !== -1) {
    var firstDebitCell = mayorSheet.getRange(firstDebitRow, 1);
    firstDebitCell.setBorder(true, false, false, true, false, false);

    var firstCreditCell = mayorSheet.getRange(firstDebitRow, 2);
    firstCreditCell.setBorder(true, true, false, false, false, false);
  }

  return startRow;
}
