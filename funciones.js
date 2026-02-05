// === WEBAPP ===
function doGet(e) {
  // Verificar si es petición al dashboard admin
  if (e && e.parameter && e.parameter.page === 'admin') {
    // Verificar que el usuario tenga permisos (debe ser el propietario o editor del script)
    const userEmail = Session.getEffectiveUser().getEmail();
    Logger.log('Usuario accediendo al admin: ' + userEmail);
    
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('Dashboard Admin - CPEM 25')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // Por defecto, mostrar la webapp pública
  return HtmlService.createTemplateFromFile('web')
    .evaluate()
    .setTitle('Sistema de Licencias CPEM 25')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerDatosHtml(nombre) {
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}
// === CONFIGURACIÓN ===
// IMPORTANTE: ID del Google Sheet de agentes
const SHEET_AGENTES_ID = '19Ik0mKdZwN37NZLBV-QLS5JLlBkUHcf3oBiIYr8u8Vo';
const SHEET_AGENTES_NOMBRE = 'RESPUESTAS'; // Nombre de la hoja dentro del archivo

// Hoja Maestra de cursos/cargos
const SHEET_MAESTRA_ID = '1B-TyiGskMJ-TDv2ck4cNyjPFMCJE6p87I3USF2zeFWc';
const SHEET_MAESTRA_NOMBRE = 'Maestra';

const SHEET_SOLICITUDES = 'SOLICITUDES';
const SHEET_JUSTIFICACIONES = 'JUSTIFICACIONES';
const SHEET_SISO = 'SISO';
const SHEET_NOVEDADES = 'NOVEDADES';
const FOLDER_ARCHIVOS_ID = '1-48O-hpqqbINKbqdvZPVLT0jjzTCOEBz'; // ID de la carpeta en Drive

// ID de la plantilla de Gmail para emails
const PLANTILLA_SOLICITUD = 'r-6079908409768349520';
const PLANTILLA_JUSTIFICACION = 'r8871403211407668186';
const PLANTILLA_ADMIN = 'r-2767467785340734120'; // Reemplazar con ID real del borrador

// ========================================
// === FUNCIONES PRINCIPALES ===
// ========================================

// === NOVEDADES (Notificaciones) ===
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

function obtenerLunesSemanaActual_() {
  const ahora = new Date();
  const dia = ahora.getDay(); // 0 domingo, 1 lunes...
  const diff = dia === 0 ? -6 : 1 - dia; // lunes de esta semana
  const lunes = new Date(ahora);
  lunes.setDate(ahora.getDate() + diff);
  lunes.setHours(0, 0, 0, 0);
  return lunes;
}

// Mapeo de columnas de tu sheet de Agentes (índice basado en 0)
const COL_AGENTES = {
  TIMESTAMP: 0,
  EMAIL: 1,
  APELLIDOS: 2,
  NOMBRES: 3,
  DNI: 4,
  NUMERO_EMPLEADO: 5,
  TELEFONO: 6,
  EMAIL_ALT: 7,
  LUGAR_NAC: 8,
  FECHA_NAC: 9,
  ESTADO_CIVIL: 10,
  DOMICILIO: 11,
  TITULO: 12,
  TITULO_OTORGADO: 13,
  TELEFONO_ALT: 14,
  FECHA_INGRESO_CPEM: 15,
  ACEPTO: 16,
  FECHA_INGRESO_CPE: 17
};

// === OBTENER DATOS DE AGENTE (desde sheet externo) ===
// Busca por DNI o Número de Empleado y retorna el registro MÁS RECIENTE
function obtenerAgente(dniONumEmpleado) {
  try {
    const ssExterno = SpreadsheetApp.openById(SHEET_AGENTES_ID);
    const sheet = ssExterno.getSheetByName(SHEET_AGENTES_NOMBRE);
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "' + SHEET_AGENTES_NOMBRE + '" en el archivo externo');
    }
    
    const data = sheet.getDataRange().getValues();
    let agentesEncontrados = [];
    
    // Buscar todos los registros que coincidan
    for (let i = 1; i < data.length; i++) {
      const dni = data[i][COL_AGENTES.DNI] ? data[i][COL_AGENTES.DNI].toString().trim() : '';
      const numEmpleado = data[i][COL_AGENTES.NUMERO_EMPLEADO] ? data[i][COL_AGENTES.NUMERO_EMPLEADO].toString().trim() : '';
      const busqueda = dniONumEmpleado.toString().trim();
      
      if (dni === busqueda || numEmpleado === busqueda) {
        agentesEncontrados.push({
          timestamp: data[i][COL_AGENTES.TIMESTAMP],
          nombre: `${data[i][COL_AGENTES.APELLIDOS]} ${data[i][COL_AGENTES.NOMBRES]}`,
          apellidos: data[i][COL_AGENTES.APELLIDOS],
          nombres: data[i][COL_AGENTES.NOMBRES],
          email: data[i][COL_AGENTES.EMAIL],
          dni: data[i][COL_AGENTES.DNI],
          numeroEmpleado: data[i][COL_AGENTES.NUMERO_EMPLEADO],
          telefono: data[i][COL_AGENTES.TELEFONO]
        });
      }
    }
    
    if (agentesEncontrados.length === 0) {
      return null;
    }
    
    // Ordenar por timestamp DESC y retornar el más reciente
    agentesEncontrados.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return agentesEncontrados[0];
    
  } catch (error) {
    Logger.log('Error al obtener agente: ' + error.toString());
    throw new Error('Error al acceder a la hoja de agentes: ' + error.toString());
  }
}

// === OBTENER CURSOS/CARGOS DEL AGENTE (desde hoja Maestra) ===
function obtenerCargosAgente(dniONumEmpleado) {
  try {

    // Obtener primero el agente para recuperar su número de empleado
    const agente = obtenerAgente(dniONumEmpleado);

    if (!agente) {
      return { success: false, error: 'Agente no encontrado en la base de datos' };
    }

    const numeroEmpleado = agente.numeroEmpleado ? agente.numeroEmpleado.toString().trim() : '';
    if (!numeroEmpleado) {
      return { success: false, error: 'El agente no tiene Número de Empleado cargado' };
    }

    const ssMaestra = SpreadsheetApp.openById(SHEET_MAESTRA_ID);
    const sheetMaestra = ssMaestra.getSheetByName(SHEET_MAESTRA_NOMBRE);

    if (!sheetMaestra) {
      return { success: false, error: 'No se encontró la hoja "' + SHEET_MAESTRA_NOMBRE + '"' };
    }

    const data = sheetMaestra.getDataRange().getValues();
    const cargos = [];

    // Buscar en columnas G (6), J (9), M (12)
    for (let i = 1; i < data.length; i++) {
      const colG = data[i][6] ? data[i][6].toString().trim() : '';
      const colJ = data[i][9] ? data[i][9].toString().trim() : '';
      const colM = data[i][12] ? data[i][12].toString().trim() : '';

      if (colG === numeroEmpleado || colJ === numeroEmpleado || colM === numeroEmpleado) {
        const colA = data[i][0] || '';
        const colB = data[i][1] || '';
        const colC = data[i][2] || '';
        const colD = data[i][3] || '';
        const colE = data[i][4] || '';
        const colF = data[i][5] || '';

        const textoMostrar = `${colC}${colD} - ${colE} - Sec: ${colA}`.trim();

        cargos.push({
          texto: textoMostrar,
          seccion: colA,
          valorOriginal: { colA, colB, colC, colD, colE, colF }
        });
      }
    }

    // Normalizar datos para evitar problemas de serialización hacia el cliente
    const agentePlano = {
      nombre: String(agente.nombre || ''),
      apellidos: String(agente.apellidos || ''),
      nombres: String(agente.nombres || ''),
      email: String(agente.email || ''),
      dni: String(agente.dni || ''),
      numeroEmpleado: String(agente.numeroEmpleado || '')
    };

    const cargosPlano = cargos.map(c => ({
      texto: String(c.texto || ''),
      seccion: String(c.seccion || '')
    }));

    return { success: true, cargos: cargosPlano, agente: agentePlano };
  } catch (error) {
    return { success: false, error: 'Error al obtener cargos: ' + error.toString() };
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
              data[i][11] === 'Injustificada')) {
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
      enviarEmailJustificacion(
        agente,
        datos.licenciasIds.length,
        archivoUrl,
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
  
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  
  // Buscar la solicitud por su ID único
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === idSolicitud) {
      sheet.getRange(i + 1, 11).setValue(nuevoEstado); // Columna K (índice 11)
      break;
    }
  }
}

// === OBTENER EMAILS DE AGENTES POR NÚMERO DE EMPLEADO ===
function obtenerEmailsAgentes(numerosEmpleado) {
  try {
    const emails = [];
    
    numerosEmpleado.forEach(numEmpleado => {
      try {
        const agente = obtenerAgente(numEmpleado);
        if (agente && agente.email) {
          emails.push(agente.email);
        }
      } catch (error) {
        Logger.log('Error al obtener email para N° ' + numEmpleado + ': ' + error.toString());
      }
    });
    
    if (emails.length === 0) {
      return { success: false, error: 'No se encontraron emails para los agentes seleccionados' };
    }
    
    return { success: true, emails: emails };
  } catch (error) {
    Logger.log('Error al obtener emails: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// === OBTENER LISTA DE AGENTES PARA JORNADA (columnas BH:BI de hoja ".") ===
function obtenerAgentesJornada() {
  try {
    const ssMaestra = SpreadsheetApp.openById(SHEET_MAESTRA_ID);
    const sheetPunto = ssMaestra.getSheetByName('.');
    
    if (!sheetPunto) {
      return { success: false, error: 'No se encontró la hoja "." en la Maestra' };
    }
    
    // BH es columna 60 (BH), BI es columna 61 (BI)
    const data = sheetPunto.getRange('BH:BI').getValues();
    const agentes = [];
    
    // Saltar el encabezado (fila 0)
    for (let i = 0; i < data.length; i++) {
      const nombre = data[i][0] ? String(data[i][0]).trim() : '';
      const numeroEmpleado = data[i][1] ? String(data[i][1]).trim() : '';
      
      // Solo agregar si ambos valores existen
      if (nombre && numeroEmpleado) {
        agentes.push({
          nombre: nombre,
          numeroEmpleado: numeroEmpleado
        });
      }
    }
    
    return { success: true, agentes: agentes };
  } catch (error) {
    Logger.log('Error al obtener agentes de jornada: ' + error.toString());
    return { success: false, error: 'Error al obtener agentes: ' + error.toString() };
  }
}

// === GENERAR SOLICITUDES MASIVAS DE JORNADA ===
function generarSolicitudesJornada(fecha, agentesSeleccionados) {
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


// === ENVIAR EMAIL - SOLICITUD (usando plantilla) ===
function enviarEmailSolicitud(agente, datos, timestamp, idSolicitud) {
  try {
    const asunto = `Solicitud de Licencia ${agente.nombre} - CPEM N° 25`;
    const draftId = PLANTILLA_SOLICITUD;
    
    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);

    const fechaCargaFormateada = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        
    // Reemplazar variables en la plantilla
    cuerpoHTML = cuerpoHTML.replace("%NOMBRE%", agente.nombre).replace("%TIPO%", datos.tipoLicencia).replace("%FECHA_CARGA%", fechaCargaFormateada);
    
    // Crear borrador con el contenido modificado
    
    GmailApp.createDraft(agente.email, asunto, "", {
      htmlBody: cuerpoHTML
    });
    
    Logger.log('Borrador de email creado exitosamente para: ' + agente.email);
  } catch (error) {
    Logger.log('Error al crear borrador de solicitud: ' + error.toString());
    throw error;
  }
}

// === ENVIAR EMAIL - JUSTIFICACIÓN (usando plantilla) ===
function enviarEmailJustificacion(agente, cantidadLicencias, archivoUrl, timestamp) {
  try {
    const asunto = "Confirmación de Licencia - CPEM N° 25";
    const draftId = PLANTILLA_JUSTIFICACION;

    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);
    
    const fechaCargaFormateada = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    
    // Reemplazar variables en la plantilla
    cuerpoHTML = cuerpoHTML.replace("%NOMBRE%", agente.nombre).replace("%CANTIDAD_LICENCIAS%", cantidadLicencias).replace("%FECHA_CARGA%", fechaCargaFormateada).replace("%ARCHIVO_URL%", archivoUrl);
    
    // Crear borrador con el contenido modificado
    GmailApp.createDraft(agente.email, asunto, "", {    
      attachments: [],
      htmlBody: cuerpoHTML
    });
    
    Logger.log('Borrador de email creado exitosamente para: ' + agente.email);
  } catch (error) {
    Logger.log('Error al crear borrador de justificación: ' + error.toString());
    throw error;
  }
}

// ========================================
// === FUNCIONES ADMINISTRATIVAS ===
// ========================================

// Obtener email del usuario actual
function obtenerEmailUsuario() {
  return Session.getEffectiveUser().getEmail();
}

// Obtener novedades desde las hojas (últimos 3 días, incluye hoy)
function obtenerNovedadesSemana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSolicitudes = ss.getSheetByName(SHEET_SOLICITUDES);
  const sheetJustificaciones = ss.getSheetByName(SHEET_JUSTIFICACIONES);

  const ahora = new Date();
  const desde = new Date(ahora);
  desde.setDate(ahora.getDate() - 3);
  desde.setHours(0, 0, 0, 0);

  const novedades = [];

  // Solicitudes
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

        novedades.push({
          timestampObj: timestamp,
          timestamp: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          tipo: 'SOLICITUD',
          mensaje: `Solicitud de licencia (${tipoLicencia}) - ${apellidos} ${nombres}`,
          dni: String(dataSol[i][2] || ''),
          numeroEmpleado: String(dataSol[i][3] || ''),
          apellidos,
          nombres,
          idsSolicitudes: id,
          origen: 'SOLICITUD'
        });
      }
    }
  }

  // Justificaciones
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

        novedades.push({
          timestampObj: timestamp,
          timestamp: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          tipo: 'JUSTIFICACION',
          mensaje: `Justificación cargada (IDs: ${ids}) - ${apellidos} ${nombres}`,
          dni: String(dataJust[i][2] || ''),
          numeroEmpleado: String(dataJust[i][3] || ''),
          apellidos,
          nombres,
          idsSolicitudes: ids,
          origen: 'JUSTIFICACION'
        });
      }
    }
  }

  // Ordenar por más reciente primero
  novedades.sort((a, b) => b.timestampObj - a.timestampObj);

  // Quitar campo interno antes de devolver
  return novedades.map(n => {
    const { timestampObj, ...rest } = n;
    return rest;
  });
}

// Obtener todas las solicitudes
function obtenerTodasSolicitudes() {
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
      id: String(data[i][12] || '')
    });
  }
  Logger.log('Solicitudes obtenidas: ' + solicitudes.length);
  if (solicitudes.length) {
    Logger.log('Primera solicitud normalizada: ' + JSON.stringify(solicitudes[0]));
  }
  
  return solicitudes;
}

  // Obtener solicitud por ID
  function obtenerSolicitudPorId(idSolicitud) {
    try {
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
            id: String(data[i][12] || '')
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

// Obtener todas las justificaciones
function obtenerTodasJustificaciones() {
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

// Actualizar estado de una fila (por índice de fila)
function actualizarEstadoFila(rowIndex, nuevoEstado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES);
    
    if (!sheet) {
      return { success: false, error: 'Hoja no encontrada' };
    }
    
    // Columna l es la 12 (Estado)
    sheet.getRange(rowIndex, 12).setValue(nuevoEstado);
    
    return { success: true };
  } catch (error) {
    Logger.log('Error al actualizar estado: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Enviar email personalizado desde admin
function enviarEmailAdmin(emailDocente, nombreDocente, mensaje) {
  try {
    const asunto = `Respuesta solicitud de licencia ${nombreDocente} - CPEM N° 25`;
    const draftId = PLANTILLA_ADMIN;
    
    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);

    // Reemplazar variables en la plantilla
    cuerpoHTML = cuerpoHTML.replace(/%RESPUESTA%/g, mensaje).replace(/%NOMBRE%/g, nombreDocente);
    
    // Crear borrador con el contenido modificado
    
    GmailApp.createDraft(emailDocente, asunto, "", {
      htmlBody: cuerpoHTML
    });
    
    return { success: true };
       
  } catch (error) {
    Logger.log('Error al crear borrador de email: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}


// === INICIALIZAR SHEETS (solo solicitudes y justificaciones) ===
function inicializarSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Solo crear las hojas nuevas, NO tocar la de Agentes
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

// === TESTING - Verificar cargos del agente ===
function testObtenerCargosAgente() {
  const dniONumEmpleado = '4017'; // Reemplazar con DNI o N° Empleado real
  // Obtener primero el agente para recuperar su número de empleado
  const agente = obtenerAgente(dniONumEmpleado);
  Logger.log('agente encontrado: ' + agente.numeroEmpleado);
  if (!agente) {
    return { success: false, error: 'Agente no encontrado en la base de datos' };
  }

  const numeroEmpleado = agente.numeroEmpleado ? agente.numeroEmpleado.toString().trim() : '';
  if (!numeroEmpleado) {
    return { success: false, error: 'El agente no tiene Número de Empleado cargado' };
  }

  try {
    const resultado = obtenerCargosAgente(numeroEmpleado);
    if (resultado.success) {
      Logger.log('✓ Cargos encontrados para el agente: ' + resultado.agente.nombre);
      Logger.log('Total de cargos: ' + resultado.cargos.length);
      resultado.cargos.forEach((cargo, index) => {
        Logger.log(`  ${index + 1}. ${cargo.texto}`);
      });
    } else {
      Logger.log('❌ Error: ' + resultado.error);
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
}


// === TESTING - Verificar conexión con sheet externo ===
function testConexionSheetExterno() {
  try {
    const ssExterno = SpreadsheetApp.openById(SHEET_AGENTES_ID);
    const sheet = ssExterno.getSheetByName(SHEET_AGENTES_NOMBRE);
    
    if (!sheet) {
      Logger.log('❌ No se encontró la hoja "' + SHEET_AGENTES_NOMBRE + '"');
      return;
    }
    
    const totalRows = sheet.getLastRow();
    Logger.log('✓ Conexión exitosa!');
    Logger.log('Total de registros (incluyendo encabezado): ' + totalRows);
    Logger.log('Encabezados: ' + sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].join(', '));
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
    Logger.log('Verifica que el ID del sheet sea correcto y tengas permisos de acceso');
  }
}

// === TESTING - Verificar búsqueda de agente ===
function testObtenerAgente() {
  const dniONumEmpleado = ' 32282931'; // Reemplazar con DNI o N° Empleado real
  try {
    const agente = obtenerAgente(dniONumEmpleado);
    if (agente) {
      Logger.log('✓ Agente encontrado:');
      Logger.log(JSON.stringify(agente, null, 2));
    } else {
      Logger.log('❌ No se encontró agente con ese DNI o N° Empleado');
    }
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
}

// ========================================
// === FUNCIONES SISO ===
// ========================================

// === GUARDAR DATOS CSV EN HOJA SISO ===
function guardarDatosCSV(datos) {
  try {
    if (!datos || datos.length < 2) {
      return { success: false, error: 'No hay datos para guardar' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_SISO);

    // Si la hoja no existe, crearla
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_SISO);
    } else {
      // Si existe, limpiar contenido previo
      sheet.clear();
    }

    // Normalizar datos: asegurar que todas las filas tengan el mismo número de columnas
    // Encontrar el número máximo de columnas
    let maxColumnas = 0;
    datos.forEach(fila => {
      if (fila.length > maxColumnas) {
        maxColumnas = fila.length;
      }
    });

    // Rellenar filas con menos columnas con valores vacíos
    const datosNormalizados = datos.map(fila => {
      const filaNormalizada = [...fila];
      while (filaNormalizada.length < maxColumnas) {
        filaNormalizada.push('');
      }
      return filaNormalizada;
    });

    // Guardar todos los datos (incluyendo encabezado)
    const numFilas = datosNormalizados.length;
    const numColumnas = maxColumnas;

    // Escribir los datos en bloque (más eficiente)
    sheet.getRange(1, 1, numFilas, numColumnas).setValues(datosNormalizados);

    // Formatear encabezado
    const rangoEncabezado = sheet.getRange(1, 1, 1, numColumnas);
    rangoEncabezado.setBackground('#4B5563');
    rangoEncabezado.setFontColor('#FFFFFF');
    rangoEncabezado.setFontWeight('bold');

    // Ajustar ancho de columnas
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