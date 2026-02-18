function asegurarHojaNovedades_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NOVEDADES);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NOVEDADES);
    sheet.appendRow([
      'Timestamp',
      'Tipo',
      'Mensaje',
      'DNI',
      'N° Empleado',
      'Apellidos',
      'Nombres',
      'IDs Solicitudes',
      'Origen'
    ]);
  }

  return sheet;
}

function registrarNovedad_(tipo, payload) {
  try {
    const sheet = asegurarHojaNovedades_();
    const timestamp = new Date();

    const mensaje = payload.mensaje || '';
    sheet.appendRow([
      timestamp,
      String(tipo || ''),
      String(mensaje || ''),
      String(payload.dni || ''),
      String(payload.numeroEmpleado || ''),
      String(payload.apellidos || ''),
      String(payload.nombres || ''),
      String(payload.idsSolicitudes || ''),
      String(payload.origen || '')
    ]);
  } catch (error) {
    Logger.log('Error al registrar novedad: ' + error.toString());
  }
}
