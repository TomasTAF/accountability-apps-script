function consolidarDatosEnMismaHoja() {
  // Abre la hoja de cálculo donde se consolidará toda la información (el archivo abierto)
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaConsolidada = spreadsheet.getSheetByName('Consolidado');
  
  // Verifica si la hoja de consolidación está vacía (sin datos)
  if (hojaConsolidada.getLastRow() < 1) {
    // Si está vacía, agrega la fila de encabezado
    hojaConsolidada.appendRow(['Fecha', 'Póliza', 'Código', 'Descripción', 'Parcial', 'Debe', 'Haber']);
  } else {
    // Si tiene datos, limpia el contenido existente excepto la fila de encabezado
    hojaConsolidada.getRange(2, 1, hojaConsolidada.getLastRow() - 1, hojaConsolidada.getLastColumn()).clearContent();
  }
  
  // Expresión regular para el patrón de nombre de las hojas
  var regex = /^\d{4}-\d{2}-\d{2} (P[IED] \d+|PD \d+)$/;
  
  // Obtiene todas las hojas dentro del documento
  var sheets = spreadsheet.getSheets();
  
  // Itera sobre cada hoja en la hoja de cálculo
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    
    // Verifica si el nombre de la hoja coincide con el patrón
    var nombreHoja = sheet.getName();
    var match = nombreHoja.match(regex);
    
    if (match) {
      // Obtiene el ID de la hoja desde el nombre de la hoja
      var idHoja = match[1];
      
      // Obtiene los datos de la hoja como una matriz
      var data = sheet.getDataRange().getValues();
      var fechaActual = '';
      
      // Procesa los datos de la tabla, excluyendo la última fila
      for (var j = 1; j < data.length - 1; j++) { // Empieza en 1 para saltar la fila de encabezados
        var fila = data[j];
        var fecha = fila[0];
        
        // Si la celda de fecha no está vacía, actualiza fechaActual
        if (fecha !== '') {
          fechaActual = fecha;
        }
        
        // Añade la fila a la hoja de consolidado con la fecha y el ID de la hoja rellenados
        hojaConsolidada.appendRow([fechaActual, idHoja, fila[1], fila[2], fila[3], fila[4], fila[5]]);
      }
    }
  }
}
