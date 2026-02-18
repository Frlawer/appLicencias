function guardarDatosCSV(datos) {
  try {
    assertAdminAutorizado_();
    if (!datos || datos.length < 2) {
      return { success: false, error: 'No hay datos para guardar' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SISO);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SISO);
    } else {
      sheet.clear();
    }

    let maxColumnas = 0;
    datos.forEach(fila => {
      if (fila.length > maxColumnas) {
        maxColumnas = fila.length;
      }
    });

    const datosNormalizados = datos.map(fila => {
      const filaNormalizada = [...fila];
      while (filaNormalizada.length < maxColumnas) {
        filaNormalizada.push('');
      }
      return filaNormalizada;
    });

    const numFilas = datosNormalizados.length;
    const numColumnas = maxColumnas;

    sheet.getRange(1, 1, numFilas, numColumnas).setValues(datosNormalizados);

    const rangoEncabezado = sheet.getRange(1, 1, 1, numColumnas);
    rangoEncabezado.setBackground('#4B5563');
    rangoEncabezado.setFontColor('#FFFFFF');
    rangoEncabezado.setFontWeight('bold');

    for (let i = 1; i <= numColumnas; i++) {
      sheet.autoResizeColumn(i);
    }

    Logger.log('Datos CSV guardados en SISO: ' + (numFilas - 1) + ' registros');

    return {
      success: true,
      registros: numFilas - 1,
      columnas: numColumnas
    };
  } catch (error) {
    Logger.log('Error al guardar datos CSV: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
