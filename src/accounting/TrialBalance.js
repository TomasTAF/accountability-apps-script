function procesarTodasLasCuentasConTotalizacion() {
  // Abrir la hoja de cálculo activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet();
  var libroMayor = hoja.getSheetByName("Libro Mayor");
  var balanza = hoja.getSheetByName("Balanza de Comprobación") || hoja.insertSheet("Balanza de Comprobación");

  // Limpiar las celdas específicas en la hoja de Balanza de Comprobación
  balanza.clear();
  // Añadir los encabezados en la primera fila
  var encabezados = ["Cuenta", "Inicial Debe", "Inicial Haber", "Póliza", "Movimientos Debe", "Movimientos Haber", "Saldo Debe", "Saldo Haber"];
  balanza.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);

  var rangoDatos = libroMayor.getDataRange();
  var datos = rangoDatos.getValues();

  var i = 0;
  var fila = 2; // Empezar en la segunda fila para la primera cuenta
  var totalInicialDebe = 0;
  var totalInicialHaber = 0;
  var totalMovimientosDebe = 0;
  var totalMovimientosHaber = 0;
  var totalSaldoDebe = 0;
  var totalSaldoHaber = 0;

  while (i < datos.length) {
    if (datos[i][0] && datos[i][0].toString().startsWith("Cuenta")) {
      var cuenta = datos[i][0].split(':')[1].trim();
      var inicialDebe = parseFloat(datos[i + 1][0]) || 0;
      var inicialHaber = parseFloat(datos[i + 1][1]) || 0;
      var poliza = datos[i + 1][2] || ""; // Obtener la póliza de la tercera columna

      // Encontrar la penúltima fila de la tabla actual
      var j = i;
      while (datos[j] && datos[j].some(cell => cell !== "")) {
        j++;
      }
      var totalDebe = parseFloat(datos[j - 2][0]) || 0;
      var totalHaber = parseFloat(datos[j - 2][1]) || 0;

      // Calcular movimientos del debe y haber
      var movimientosDebe = totalDebe - inicialDebe;
      var movimientosHaber = totalHaber - inicialHaber;

      // Obtener el saldo de la última fila de la tabla
      var saldoDebe = parseFloat(datos[j - 1][0]) || 0;
      var saldoHaber = parseFloat(datos[j - 1][1]) || 0;

      // Determinar en cuál columna poner el saldo y asegurarse de que sean valores absolutos
      var saldoDebeFinal = saldoDebe !== 0 ? saldoDebe : 0;
      var saldoHaberFinal = saldoHaber !== 0 ? Math.abs(saldoHaber) : 0;

      // Colocar los valores en la hoja Balanza de Comprobación
      balanza.getRange(fila, 1).setValue(cuenta);
      balanza.getRange(fila, 2).setValue(inicialDebe);
      balanza.getRange(fila, 3).setValue(inicialHaber);
      balanza.getRange(fila, 4).setValue(poliza); // Añadir la póliza
      balanza.getRange(fila, 5).setValue(movimientosDebe);
      balanza.getRange(fila, 6).setValue(movimientosHaber);
      balanza.getRange(fila, 7).setValue(saldoDebeFinal);
      balanza.getRange(fila, 8).setValue(saldoHaberFinal);

      // Acumular totales para totalización
      totalInicialDebe += inicialDebe;
      totalInicialHaber += inicialHaber;
      totalMovimientosDebe += movimientosDebe;
      totalMovimientosHaber += movimientosHaber;
      totalSaldoDebe += saldoDebeFinal;
      totalSaldoHaber += saldoHaberFinal;

      // Imprimir el contenido de la tabla en el Log
      Logger.log("Tabla identificada:");
      var k = i;
      while (datos[k] && datos[k].some(cell => cell !== "")) {
        Logger.log(datos[k]);
        k++;
      }

      // Incrementar la fila para la siguiente cuenta
      fila++;

      // Avanzar al siguiente conjunto de datos en blanco antes de continuar con la próxima cuenta
      while (datos[i] && datos[i].some(cell => cell === "")) {
        i++;
      }
    }
    i++;
  }

  // Añadir la fila de totalización
  balanza.getRange(fila, 1).setValue("Total");
  balanza.getRange(fila, 2).setValue(totalInicialDebe);
  balanza.getRange(fila, 3).setValue(totalInicialHaber);
  balanza.getRange(fila, 4).setValue(""); // Columna de póliza vacía para la fila de totalización
  balanza.getRange(fila, 5).setValue(totalMovimientosDebe);
  balanza.getRange(fila, 6).setValue(totalMovimientosHaber);
  balanza.getRange(fila, 7).setValue(totalSaldoDebe);
  balanza.getRange(fila, 8).setValue(totalSaldoHaber);
}
