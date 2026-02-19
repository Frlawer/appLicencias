// === WEBAPP ===
function doGet(e) {
  // Verificar si es petición al dashboard admin
  if (e && e.parameter && e.parameter.page === 'admin') {
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log('Usuario accediendo al admin: ' + userEmail);

    if (!esAdminAutorizado_()) {
      return HtmlService.createHtmlOutput('<h2>Acceso denegado</h2><p>No tenés permisos para acceder al panel de administración.</p>')
        .setTitle('Acceso denegado');
    }
    
    // Usar createTemplateFromFile para procesar los <?!= ?> tags
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('Dashboard Admin - CPEM 25')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
  
  // Por defecto, mostrar la webapp pública
  return HtmlService.createTemplateFromFile('web')
    .evaluate()
    .setTitle('Sistema de Licencias CPEM 25')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerDatosHtml(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}

// ========================================
// === FUNCIONES PRINCIPALES ===
// ========================================


function obtenerLunesSemanaActual_() {
  const ahora = new Date();
  const dia = ahora.getDay(); // 0 domingo, 1 lunes...
  const diff = dia === 0 ? -6 : 1 - dia; // lunes de esta semana
  const lunes = new Date(ahora);
  lunes.setDate(ahora.getDate() + diff);
  lunes.setHours(0, 0, 0, 0);
  return lunes;
}


// === OBTENER LICENCIAS (desde hoja LICENCIAS) ===
// Lee la Sheet "LICENCIAS" y retorna array con: nombre, dias_totales, tipo_dias
function obtenerLicencias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LICENCIAS);
    
    if (!sheet) {
      Logger.log('Advertencia: No se encontró la hoja "' + SHEET_LICENCIAS + '"');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const licencias = [];
    
    // Saltar encabezado (fila 0)
    // Columna A (0): nombre licencia
    // Columna B (1): dias_totales
    // Columna C (2): tipo_dias (habiles/continuos)
    // Columna D (3): activa (SI/NO)
    for (let i = 1; i < data.length; i++) {
      const nombre = data[i][0] ? data[i][0].toString().trim() : '';
      const diasTotales = data[i][1];
      const tipoDias = data[i][2] ? data[i][2].toString().trim().toLowerCase() : 'continuos';
      const activa = data[i][3] ? data[i][3].toString().trim().toUpperCase() : 'SI';
      
      // Solo incluir si está activa y tiene nombre
      if (nombre && activa === 'SI') {
        licencias.push({
          nombre: nombre,
          dias_totales: diasTotales,
          tipo_dias: tipoDias
        });
      }
    }
    
    return licencias;
  } catch (error) {
    Logger.log('Error al obtener licencias: ' + error.toString());
    return [];
  }
}


// === GUARDAR SOLICITUD ===
function guardarSolicitud(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SOLICITUDES);
      sheet.appendRow([
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
    
    // Buscar agente por DNI o N° Empleado
    const agente = obtenerAgente(datos.dniONumEmpleado);
    if (!agente) {
      return { success: false, error: 'DNI o Número de Empleado no encontrado en la base de datos' };
    }
    
    const timestamp = new Date();
    const id = Utilities.getUuid();
    
    sheet.appendRow([
      timestamp,
      agente.email,
      agente.dni,
      agente.numeroEmpleado || '',
      agente.apellidos,
      agente.nombres,
      datos.fechaDesde,
      datos.fechaHasta,
      datos.cursoOCargo || '',
      datos.articulacion,
      datos.tipoLicencia,
      'Pendiente',
      id
    ]);

    registrarNovedad_('SOLICITUD', {
      mensaje: `Solicitud de licencia (${datos.tipoLicencia}) cargada por ${agente.apellidos} ${agente.nombres}`,
      dni: agente.dni,
      numeroEmpleado: agente.numeroEmpleado || '',
      apellidos: agente.apellidos,
      nombres: agente.nombres,
      idsSolicitudes: id,
      origen: 'SOLICITUD'
    });
    
    // Enviar email al agente
    try {
      enviarEmailSolicitud(agente, datos, timestamp, id);
    } catch (emailError) {
      Logger.log('Error al enviar email: ' + emailError.toString());
      // No fallar la operación si el email falla
      
    }

    // Normalizar datos para evitar problemas de serialización hacia el cliente
    const agentePlano = {
      nombre: agente.nombre || '',
      email: agente.email || '',
      dni: agente.dni || '',
      numeroEmpleado: agente.numeroEmpleado || ''
    };

    return { success: true, id: id, agente: agentePlano };

  } catch (error) {

    return { success: false, error: error.toString() };
    
  }
}

// === OBTENER SOLICITUDES DE UN AGENTE ===
function obtenerSolicitudesAgente(dniONumEmpleado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
  
  if (!sheet) return [];
  
  // Primero obtener datos del agente
  const agente = obtenerAgente(dniONumEmpleado);
  if (!agente) return [];
  
  const data = sheet.getDataRange().getValues();
  const solicitudes = [];
  
  // Buscar por DNI o N° Empleado
  for (let i = 1; i < data.length; i++) {
    const dniSheet = data[i][2] ? data[i][2].toString().trim() : '';
    const numEmpSheet = data[i][3] ? data[i][3].toString().trim() : '';
    
    if ((dniSheet === agente.dni.toString() || numEmpSheet === agente.numeroEmpleado.toString()) 
        && (
              data[i][11] === 'Autorizada' || 
              data[i][11] === 'Pendiente' ||
              data[i][11] === 'Completada' ||
              data[i][11] === 'Justificada' ||
              data[i][11] === 'Injustificada' ||
              data[i][11] === 'Descuento')) {
      solicitudes.push({
        id: data[i][12],
        rowIndex: i + 1,
        tipo: data[i][10],
        estado: data[i][11],
        desde: Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hasta: Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        cursoOCargo: data[i][8]
      });
    }
  }
  
  return solicitudes;
}

// === GUARDAR JUSTIFICACIÓN (múltiples archivos) ===
function guardarJustificacion(datos, archivos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_JUSTIFICACIONES);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_JUSTIFICACIONES);
      sheet.appendRow([
        "Timestamp",
        "Email",
        "DNI",
        "N° Empleado",
        "Apellidos",
        "Nombres",
        "IDs Solicitudes",
        "Cantidad Licencias",
        "URL Archivo",
      ]);
    }

    // Buscar agente por DNI o N° Empleado
    const agente = obtenerAgente(datos.dniONumEmpleado);
    if (!agente) {
      return {
        success: false,
        error: "DNI o Número de Empleado no encontrado en la base de datos",
      };
    }

    // Obtener carpeta por ID
    const folder = DriveApp.getFolderById(FOLDER_ARCHIVOS_ID);

    // Guardar archivos en Drive
    const urls = [];
    if (archivos && archivos.length) {
      archivos.forEach((archivo, idx) => {
        try {
          if (archivo.archivoBase64 && archivo.archivoBase64.includes(",")) {
            const contentType = archivo.mimeType || "application/pdf";
            const extension = contentType === "application/pdf" ? "pdf" : "jpg";
            const nombreBase = `${agente.dni}-${datos.licenciasIds.join("_")}-${idx + 1}`;
            const blob = Utilities.newBlob(
              Utilities.base64Decode(archivo.archivoBase64.split(",")[1]),
              contentType,
              `${nombreBase}.${extension}`
            );
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            urls.push(file.getUrl());
          }
        } catch (err) {
          Logger.log('Error al guardar archivo: ' + err.toString());
        }
      });
    }

    const timestamp = new Date();
    const idsString = datos.licenciasIds.join(", ");

    sheet.appendRow([
      timestamp,
      agente.email,
      agente.dni,
      agente.numeroEmpleado,
      agente.apellidos,
      agente.nombres,
      idsString,
      datos.licenciasIds.length,
      urls.join(", "),
    ]);

    registrarNovedad_('JUSTIFICACION', {
      mensaje: `Justificación cargada por ${agente.apellidos} ${agente.nombres} (IDs: ${idsString})`,
      dni: agente.dni,
      numeroEmpleado: agente.numeroEmpleado || '',
      apellidos: agente.apellidos,
      nombres: agente.nombres,
      idsSolicitudes: idsString,
      origen: 'JUSTIFICACION'
    });

    // Actualizar estado de todas las solicitudes seleccionadas
    datos.licenciasIds.forEach((id) => {
      actualizarEstadoSolicitud(id, "Justificada");
    });

    // Enviar email al agente
    try {
      const urlsTexto = urls.join(', ');
      enviarEmailJustificacion(
        agente,
        datos.licenciasIds.length,
        urlsTexto,
        timestamp
      );
    } catch (emailError) {
      Logger.log("Error al enviar email: " + emailError.toString());
      // No fallar la operación si el email falla
    }

    // Normalizar datos para evitar problemas de serialización hacia el cliente
    const agentePlano = {
      email: agente.email || "",
    };

    return { success: true, email: agentePlano.email };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// === ACTUALIZAR ESTADO DE SOLICITUD ===
function actualizarEstadoSolicitud(idSolicitud, nuevoEstado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
  
  if (!sheet) return { success: false, error: 'Hoja no encontrada' };
  
  const data = sheet.getDataRange().getValues();
  
  // Buscar la solicitud por su ID único
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][12] || '') === String(idSolicitud)) {
      sheet.getRange(i + 1, 12).setValue(nuevoEstado); // Columna L (Estado)
      return { success: true };
    }
  }

  return { success: false, error: 'No se encontró solicitud con el ID indicado' };
}

// === GENERAR SOLICITUDES MASIVAS DE JORNADA ===
function generarSolicitudesJornada(fecha, agentesSeleccionados) {
  try {
    assertAdminAutorizado_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SOLICITUDES);
      sheet.appendRow([
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
    
    const timestamp = new Date();
    let exitosas = 0;
    let fallidas = 0;
    const errores = [];
    
    agentesSeleccionados.forEach(numEmpleado => {
      try {
        // Buscar agente por número de empleado
        const agente = obtenerAgente(numEmpleado);
        
        if (!agente) {
          fallidas++;
          errores.push(`N° Empleado ${numEmpleado}: No encontrado en base de datos`);
          return;
        }
        
        const id = Utilities.getUuid();
        
        sheet.appendRow([
          timestamp,
          agente.email || '',
          agente.dni || '',
          agente.numeroEmpleado || '',
          agente.apellidos || '',
          agente.nombres || '',
          fecha,
          fecha,
          'Todos',
          'No',
          'JORNADA INSTITUCIONAL',
          'Autorizada',
          id
        ]);
        
        exitosas++;
      } catch (error) {
        fallidas++;
        errores.push(`N° Empleado ${numEmpleado}: ${error.toString()}`);
      }
    });
    
    return {
      success: true,
      exitosas: exitosas,
      fallidas: fallidas,
      errores: errores
    };
    
  } catch (error) {
    Logger.log('Error al generar solicitudes de jornada: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}



// Funciones administrativas movidas a admin.service.gs

