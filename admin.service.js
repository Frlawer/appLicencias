function obtenerEmailUsuario() {
  assertAdminAutorizado_();
  return Session.getEffectiveUser().getEmail();
}

function obtenerNovedadesSemana() {
  assertAdminAutorizado_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSolicitudes = ss.getSheetByName(SHEET_SOLICITUDES);
  const sheetJustificaciones = ss.getSheetByName(SHEET_JUSTIFICACIONES);

  const ahora = new Date();
  const desde = new Date(ahora);
  desde.setDate(ahora.getDate() - 3);
  desde.setHours(0, 0, 0, 0);

  const novedades = [];
  const tiposPorId = {};

  if (sheetSolicitudes) {
    const dataSol = sheetSolicitudes.getDataRange().getValues();
    for (let i = 1; i < dataSol.length; i++) {
      const ts = dataSol[i][0];
      const timestamp = ts instanceof Date ? ts : new Date(ts);
      if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) continue;

      if (timestamp >= desde) {
        const apellidos = String(dataSol[i][4] || '');
        const nombres = String(dataSol[i][5] || '');
        const tipoLicencia = String(dataSol[i][10] || '');
        const id = String(dataSol[i][12] || '');

        if (id) {
          tiposPorId[id] = tipoLicencia;
        }

        novedades.push({
          timestampObj: timestamp,
          timestamp: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          tipo: 'SOLICITUD',
          mensaje: `Solicitud de licencia (${tipoLicencia}) - ${apellidos} ${nombres}`,
          dni: String(dataSol[i][2] || ''),
          numeroEmpleado: String(dataSol[i][3] || ''),
          apellidos,
          nombres,
          tipoLicencia,
          idsSolicitudes: id,
          origen: 'SOLICITUD'
        });
      }
    }
  }

  if (sheetJustificaciones) {
    const dataJust = sheetJustificaciones.getDataRange().getValues();
    for (let i = 1; i < dataJust.length; i++) {
      const ts = dataJust[i][0];
      const timestamp = ts instanceof Date ? ts : new Date(ts);
      if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) continue;

      if (timestamp >= desde) {
        const apellidos = String(dataJust[i][4] || '');
        const nombres = String(dataJust[i][5] || '');
        const ids = String(dataJust[i][6] || '');
        const idsSeparados = ids
          .split(',')
          .map(id => String(id || '').trim())
          .filter(Boolean);
        const tiposRelacionados = idsSeparados
          .map(id => tiposPorId[id])
          .filter(Boolean);
        const tiposUnicos = [...new Set(tiposRelacionados)];
        const tipoLicencia = tiposUnicos.join(', ');

        novedades.push({
          timestampObj: timestamp,
          timestamp: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          tipo: 'JUSTIFICACION',
          mensaje: `Justificación cargada (IDs: ${ids}) - ${apellidos} ${nombres}`,
          dni: String(dataJust[i][2] || ''),
          numeroEmpleado: String(dataJust[i][3] || ''),
          apellidos,
          nombres,
          tipoLicencia,
          idsSolicitudes: ids,
          origen: 'JUSTIFICACION'
        });
      }
    }
  }

  novedades.sort((a, b) => b.timestampObj - a.timestampObj);

  return novedades.map(n => {
    const { timestampObj, ...rest } = n;
    return rest;
  });
}

function obtenerTodasSolicitudes() {
  assertAdminAutorizado_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SOLICITUDES);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const solicitudes = [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : (data[i][0] ? String(data[i][0]) : '');
    const fechaDesde = data[i][6] instanceof Date
      ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (data[i][6] ? String(data[i][6]) : '');
    const fechaHasta = data[i][7] instanceof Date
      ? Utilities.formatDate(data[i][7], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : (data[i][7] ? String(data[i][7]) : '');

    solicitudes.push({
      rowIndex: Number(i + 1),
      timestamp: timestamp,
      email: String(data[i][1] || ''),
      dni: String(data[i][2] || ''),
      numeroEmpleado: String(data[i][3] || ''),
      apellidos: String(data[i][4] || ''),
      nombres: String(data[i][5] || ''),
      fechaDesde: fechaDesde,
      fechaHasta: fechaHasta,
      cursoOCargo: String(data[i][8] || ''),
      articulacion: String(data[i][9] || ''),
      tipoLicencia: String(data[i][10] || ''),
      estado: String(data[i][11] || ''),
      id: String(data[i][12] || ''),
      motivo: String(data[i][13] || '')
    });
  }
  Logger.log('Solicitudes obtenidas: ' + solicitudes.length);
  if (solicitudes.length) {
    Logger.log('Primera solicitud normalizada: ' + JSON.stringify(solicitudes[0]));
  }

  return solicitudes;
}

function obtenerSolicitudPorId(idSolicitud) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    if (!sheet) {
      return { success: false, error: 'Hoja de solicitudes no encontrada' };
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][12] || '') === String(idSolicitud)) {
        const timestamp = data[i][0] instanceof Date
          ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
          : (data[i][0] ? String(data[i][0]) : '');
        const fechaDesde = data[i][6] instanceof Date
          ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : (data[i][6] ? String(data[i][6]) : '');
        const fechaHasta = data[i][7] instanceof Date
          ? Utilities.formatDate(data[i][7], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : (data[i][7] ? String(data[i][7]) : '');

        const solicitud = {
          rowIndex: Number(i + 1),
          timestamp,
          email: String(data[i][1] || ''),
          dni: String(data[i][2] || ''),
          numeroEmpleado: String(data[i][3] || ''),
          apellidos: String(data[i][4] || ''),
          nombres: String(data[i][5] || ''),
          fechaDesde,
          fechaHasta,
          cursoOCargo: String(data[i][8] || ''),
          articulacion: String(data[i][9] || ''),
          tipoLicencia: String(data[i][10] || ''),
          estado: String(data[i][11] || ''),
          id: String(data[i][12] || ''),
          motivo: String(data[i][13] || '')
        };
        return { success: true, solicitud };
      }
    }

    return { success: false, error: 'No se encontró la solicitud con ese ID' };
  } catch (error) {
    Logger.log('Error al obtener solicitud por ID: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function obtenerTodasJustificaciones() {
  assertAdminAutorizado_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_JUSTIFICACIONES);

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const justificaciones = [];

  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0] instanceof Date
      ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : (data[i][0] ? String(data[i][0]) : '');

    justificaciones.push({
      rowIndex: Number(i + 1),
      timestamp: timestamp,
      email: String(data[i][1] || ''),
      dni: String(data[i][2] || ''),
      numeroEmpleado: String(data[i][3] || ''),
      apellidos: String(data[i][4] || ''),
      nombres: String(data[i][5] || ''),
      idsSolicitudes: String(data[i][6] || ''),
      cantidadLicencias: Number(data[i][7]) || 0,
      archivoUrl: String(data[i][8] || '')
    });
  }

  return justificaciones;
}

function actualizarEstadoFila(rowIndex, nuevoEstado) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);

    if (!sheet) {
      return { success: false, error: 'Hoja no encontrada' };
    }

    sheet.getRange(rowIndex, 12).setValue(nuevoEstado);

    return { success: true };
  } catch (error) {
    Logger.log('Error al actualizar estado: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function actualizarEstadosFilas(rowIndexes, nuevoEstado) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);

    if (!sheet) {
      return { success: false, error: 'Hoja no encontrada' };
    }

    const registros = Array.isArray(rowIndexes) ? rowIndexes : [rowIndexes];
    const validos = registros
      .map(Number)
      .filter(n => Number.isFinite(n) && n >= 2 && n <= sheet.getLastRow());

    if (!validos.length) {
      return { success: false, error: 'No hay filas válidas para actualizar' };
    }

    validos.forEach(rowIndex => {
      sheet.getRange(rowIndex, 12).setValue(nuevoEstado);
    });

    return { success: true, actualizados: validos.length };
  } catch (error) {
    Logger.log('Error al actualizar estados masivos: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function eliminarSolicitud(rowIndex) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);

    if (!sheet) {
      return { success: false, error: 'Hoja no encontrada' };
    }

    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return { success: false, error: 'Índice de fila inválido' };
    }

    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (error) {
    Logger.log('Error al eliminar solicitud: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function eliminarSolicitudes(rowIndexes) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);

    if (!sheet) {
      return { success: false, error: 'Hoja no encontrada' };
    }

    const registros = Array.isArray(rowIndexes) ? rowIndexes : [rowIndexes];
    const validos = registros
      .map(Number)
      .filter(n => Number.isFinite(n) && n >= 2 && n <= sheet.getLastRow())
      .sort((a, b) => b - a);

    if (!validos.length) {
      return { success: false, error: 'No hay solicitudes válidas para eliminar' };
    }

    validos.forEach(rowIndex => {
      sheet.deleteRow(rowIndex);
    });

    return { success: true, eliminadas: validos.length };
  } catch (error) {
    Logger.log('Error al eliminar solicitudes masivas: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function extraerIdsArchivosDrive_(archivoUrl) {
  const urls = String(archivoUrl || '')
    .split(',')
    .map(u => u.trim())
    .filter(Boolean);

  const ids = new Set();

  urls.forEach(url => {
    let id = '';
    const matchPath = url.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
    const matchQuery = url.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);

    if (matchPath && matchPath[1]) {
      id = matchPath[1];
    } else if (matchQuery && matchQuery[1]) {
      id = matchQuery[1];
    } else {
      const fallback = url.match(/[a-zA-Z0-9_-]{25,}/);
      if (fallback && fallback[0]) id = fallback[0];
    }

    if (id) ids.add(id);
  });

  return Array.from(ids);
}

function eliminarJustificacion(rowIndex) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_JUSTIFICACIONES);

    if (!sheet) {
      return { success: false, error: 'Hoja de justificaciones no encontrada' };
    }

    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      return { success: false, error: 'Índice de fila inválido' };
    }

    const archivoUrl = String(sheet.getRange(rowIndex, 9).getValue() || '');
    const fileIds = extraerIdsArchivosDrive_(archivoUrl);

    for (let i = 0; i < fileIds.length; i++) {
      const fileId = fileIds[i];
      try {
        DriveApp.getFileById(fileId);
      } catch (errorFile) {
        return { success: false, error: 'No se pudo acceder al archivo de Drive: ' + fileId };
      }
    }

    fileIds.forEach(fileId => {
      DriveApp.getFileById(fileId).setTrashed(true);
    });

    sheet.deleteRow(rowIndex);

    return {
      success: true,
      archivosEliminados: fileIds.length
    };
  } catch (error) {
    Logger.log('Error al eliminar justificación: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function inicializarSheets() {
  assertAdminAutorizado_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetSolicitudes = ss.getSheetByName(SHEET_SOLICITUDES);
  if (!sheetSolicitudes) {
    sheetSolicitudes = ss.insertSheet(SHEET_SOLICITUDES);
    sheetSolicitudes.appendRow([
      'Timestamp',
      'Email',
      'DNI',
      'N° Empleado',
      'Apellidos',
      'Nombres',
      'Fecha Desde',
      'Fecha Hasta',
      'Curso/Cargo',
      'Articulación',
      'Tipo Licencia',
      'Estado',
      'ID'
    ]);
  }

  let sheetJustificaciones = ss.getSheetByName(SHEET_JUSTIFICACIONES);
  if (!sheetJustificaciones) {
    sheetJustificaciones = ss.insertSheet(SHEET_JUSTIFICACIONES);
    sheetJustificaciones.appendRow([
      'Timestamp',
      'Email',
      'DNI',
      'N° Empleado',
      'Apellidos',
      'Nombres',
      'IDs Solicitudes',
      'Cantidad Licencias',
      'URL Archivo'
    ]);
  }

  Logger.log('Sheets inicializados correctamente');
}

// ========================================
// === FUNCIONES RAZONES PARTICULARES ===
// ========================================

function obtenerSolicitudesRazonesPart() {
  try {
    assertAdminAutorizado_();
    const todasSolicitudes = obtenerTodasSolicitudes();
    
    // Filtrar solo solicitudes que contengan "*" en el tipo de licencia (RAZONES PARTICULARES)
    const rpSolicitudes = todasSolicitudes.filter(sol => {
      return sol.tipoLicencia && sol.tipoLicencia.includes('*');
    });

    // Ordenar: más recientes primero
    rpSolicitudes.sort((a, b) => {
      const fechaA = new Date(a.timestamp || 0).getTime();
      const fechaB = new Date(b.timestamp || 0).getTime();
      if (fechaB !== fechaA) return fechaB - fechaA;
      return Number(b.rowIndex || 0) - Number(a.rowIndex || 0);
    });

    return rpSolicitudes;
  } catch (error) {
    Logger.log('Error al obtener solicitudes de RP: ' + error.toString());
    return [];
  }
}

function obtenerAgentesConLimiteMensualRP() {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRP = ss.getSheetByName('RP');

    if (!sheetRP || sheetRP.getLastRow() < 2) {
      return [];
    }

    const values = sheetRP.getRange(2, 1, sheetRP.getLastRow() - 1, 9).getValues();
    const vistos = {};
    const resultado = [];

    for (const row of values) {
      const agente = String(row[1] || '').trim();
      const cargo = String(row[2] || '').trim();
      const limiteMensual = String(row[8] || '').trim();
      if (!agente || !cargo || !limiteMensual.includes('✗')) continue;

      const clave = `${agente}|${cargo}`;
      if (vistos[clave]) continue;
      vistos[clave] = true;

      resultado.push({
        agente,
        cargo,
        limiteMensual
      });
    }

    return resultado;
  } catch (error) {
    Logger.log('Error al obtener agentes con límite mensual RP: ' + error.toString());
    return [];
  }
}

function generarReporteRP() {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Crear o limpiar la hoja "RP"
    let sheetRP = ss.getSheetByName('RP');
    if (!sheetRP) {
      sheetRP = ss.insertSheet('RP');
    } else {
      // Limpiar la hoja excepto el encabezado
      const lastRow = sheetRP.getLastRow();
      if (lastRow > 1) {
        sheetRP.deleteRows(2, lastRow - 1);
      }
    }

    // Crear encabezados si la hoja está vacía
    if (sheetRP.getLastRow() === 0) {
      sheetRP.appendRow([
        'N° Empleado',
        'Agente',
        'Cargo',
        'Curso',
        'División',
        'Solicitudes Anuales',
        'Solicitudes Mensuales',
        'Límite Anual (6)',
        'Límite Mensual (2)'
      ]);
      // Dar formato al encabezado
      const headerRange = sheetRP.getRange(1, 1, 1, 9);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2B3E4C');
      headerRange.setFontColor('#FFFFFF');
    }

    // Obtener todas las solicitudes de RP
    const solicitudesRP = obtenerSolicitudesRazonesPart();
    
    if (solicitudesRP.length === 0) {
      return { success: true, mensaje: 'No hay solicitudes de Razones Particulares registradas', registros: 0 };
    }

    // Agrupar por agente + cargo
    const reportePorAgenteYCargo = {};
    
    for (const solicitud of solicitudesRP) {
      // Obtener datos del agente desde la clave maestra
      const agenteData = obtenerAgente(solicitud.numeroEmpleado);
      if (!agenteData) continue;

      // Obtener cargos del agente
      const cargosData = obtenerCargosAgente(solicitud.numeroEmpleado);
      if (!cargosData.success || !cargosData.cargos.length) continue;

      // Para cada cargo, contar las solicitudes
      for (const cargo of cargosData.cargos) {
        const clave = `${solicitud.numeroEmpleado}|${cargo.texto}`;
        
        if (!reportePorAgenteYCargo[clave]) {
          reportePorAgenteYCargo[clave] = {
            numeroEmpleado: solicitud.numeroEmpleado,
            agente: agenteData.nombre,
            cargo: cargo.texto,
            seccion: cargo.seccion,
            solicitudesAño: 0,
            solicitudesMes: 0,
            solicitudes: []
          };
        }

        // Contar por año calendario
        const fechaSolicitud = new Date(solicitud.timestamp);
        const mesActual = new Date().getMonth();
        const anoActual = new Date().getFullYear();
        const anoSolicitud = fechaSolicitud.getFullYear();
        const mesSolicitud = fechaSolicitud.getMonth();

        // Solicitudes del año calendario actual
        if (anoSolicitud === anoActual) {
          reportePorAgenteYCargo[clave].solicitudesAño++;
        }

        // Solicitudes del mes actual
        if (anoSolicitud === anoActual && mesSolicitud === mesActual) {
          reportePorAgenteYCargo[clave].solicitudesMes++;
        }

        reportePorAgenteYCargo[clave].solicitudes.push({
          timestamp: solicitud.timestamp,
          motivo: solicitud.motivo,
          estado: solicitud.estado
        });
      }
    }

    // Agregar filas al reporte
    for (const clave in reportePorAgenteYCargo) {
      const item = reportePorAgenteYCargo[clave];
      
      // Extraer Curso y División de la sección (formato: "División - Curso")
      const partes = item.seccion.split('-');
      const division = partes[0] ? partes[0].trim() : '';
      const curso = partes.length > 1 ? partes[1].trim() : item.seccion;

      sheetRP.appendRow([
        item.numeroEmpleado,
        item.agente,
        item.cargo,
        curso,
        division,
        item.solicitudesAño,
        item.solicitudesMes,
        item.solicitudesAño >= 6 ? '✗ LÍMITE ALCANZADO' : item.solicitudesAño + '/6',
        item.solicitudesMes >= 2 ? '✗ LÍMITE ALCANZADO' : item.solicitudesMes + '/2'
      ]);
    }

    // Ajustar anchos de columna
    sheetRP.setColumnWidth(1, 120);
    sheetRP.setColumnWidth(2, 200);
    sheetRP.setColumnWidth(3, 250);
    sheetRP.setColumnWidth(4, 150);
    sheetRP.setColumnWidth(5, 150);
    sheetRP.setColumnWidth(6, 120);
    sheetRP.setColumnWidth(7, 120);
    sheetRP.setColumnWidth(8, 150);
    sheetRP.setColumnWidth(9, 150);

    return { 
      success: true, 
      mensaje: `Reporte generado exitosamente con ${Object.keys(reportePorAgenteYCargo).length} registros`, 
      registros: Object.keys(reportePorAgenteYCargo).length 
    };

  } catch (error) {
    Logger.log('Error al generar reporte RP: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// === GUARDAR JUSTIFICACIÓN CON ARCHIVOS EN DRIVE ===
function guardarJustificacionBackend(datos, archivos) {
  try {
    // Validar datos básicos
    if (!datos || !datos.dniONumEmpleado) {
      return { success: false, error: 'DNI o Número de Empleado requerido' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetJust = ss.getSheetByName(SHEET_JUSTIFICACIONES);

    // Crear sheet si no existe
    if (!sheetJust) {
      sheetJust = ss.insertSheet(SHEET_JUSTIFICACIONES);
      sheetJust.appendRow([
        'Timestamp',
        'Email',
        'DNI',
        'N° Empleado',
        'Apellidos',
        'Nombres',
        'IDs Solicitudes',
        'Cantidad Licencias',
        'URLs Archivos'
      ]);
    }

    // Buscar agente
    const agente = obtenerAgente(datos.dniONumEmpleado);
    if (!agente) {
      return { 
        success: false, 
        error: 'DNI o Número de Empleado no encontrado en la base de datos' 
      };
    }

    // Procesar y subir archivos a Drive
    const urls = [];
    const licenciasIds = (datos.licenciasIds && Array.isArray(datos.licenciasIds)) 
      ? datos.licenciasIds 
      : (datos.licencias && Array.isArray(datos.licencias) 
          ? datos.licencias.map(l => l.id || l) 
          : []);

    if (archivos && Array.isArray(archivos) && archivos.length > 0) {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ARCHIVOS_ID);
        
        archivos.forEach((archivoData, idx) => {
          try {
            if (!archivoData.archivoBase64) {
              Logger.log(`Archivo ${idx + 1} sin contenido base64`);
              return;
            }

            // Separar data URL si existe
            const dataParte = archivoData.archivoBase64.includes(',') 
              ? archivoData.archivoBase64.split(',')[1] 
              : archivoData.archivoBase64;

            const mimeType = archivoData.mimeType || 'application/octet-stream';
            const extension = mimeType.includes('pdf') ? 'pdf' : 'jpg';
            const timestamp = new Date().getTime();
            const nombreArchivo = `${agente.dni}_${licenciasIds.join('-')}_${timestamp}_${idx + 1}.${extension}`;

            // Decodificar y crear archivo en Drive
            const blob = Utilities.newBlob(
              Utilities.base64Decode(dataParte),
              mimeType,
              nombreArchivo
            );
            
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            const fileUrl = file.getUrl();
            urls.push(fileUrl);

            Logger.log(`Archivo subido: ${nombreArchivo} - ${fileUrl}`);
          } catch (fileError) {
            Logger.log(`Error al procesar archivo ${idx + 1}: ${fileError.toString()}`);
          }
        });
      } catch (driveError) {
        Logger.log(`Error al acceder a Drive: ${driveError.toString()}`);
        return { 
          success: false, 
          error: `Error al subir archivos: ${driveError.toString()}` 
        };
      }
    }

    // Guardar registro en sheet de justificaciones
    const timestamp = new Date();
    const idsString = licenciasIds.join(', ');

    try {
      sheetJust.appendRow([
        timestamp,
        agente.email,
        agente.dni,
        agente.numeroEmpleado,
        agente.apellidos,
        agente.nombres,
        idsString,
        licenciasIds.length,
        urls.join(', ', )
      ]);
    } catch (sheetError) {
      Logger.log(`Error al guardar en sheet: ${sheetError.toString()}`);
      return { 
        success: false, 
        error: `Error al guardar justificación: ${sheetError.toString()}` 
      };
    }

    // Actualizar estado de solicitudes
    const sheetSol = ss.getSheetByName(SHEET_SOLICITUDES);
    if (sheetSol && licenciasIds.length > 0) {
      const dataSol = sheetSol.getDataRange().getValues();
      for (let i = 1; i < dataSol.length; i++) {
        const idEnFila = String(dataSol[i][12] || '');
        if (licenciasIds.includes(idEnFila) || licenciasIds.find(id => String(id) === idEnFila)) {
          sheetSol.getRange(i + 1, 12).setValue('Justificada');
        }
      }
    }

    // Registrar novedad
    registrarNovedad_('JUSTIFICACION', {
      dni: agente.dni,
      numeroEmpleado: agente.numeroEmpleado || '',
      apellidos: agente.apellidos,
      nombres: agente.nombres,
      idsSolicitudes: idsString,
      origen: 'JUSTIFICACION'
    });

    return { 
      success: true, 
      email: agente.email,
      archivosSubidos: urls.length,
      mensaje: `Justificación guardada con ${urls.length} archivo(s)` 
    };

  } catch (error) {
    Logger.log('Error en guardarJustificacionBackend: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}


