// === WEBAPP ===
function doGet()
{
  return HtmlService.createTemplateFromFile('web').evaluate().setTitle('Sistema de Licencias CPEM 25').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerDatosHtml(nombre)
{
  return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}
// === CONFIGURACIÓN ===
// IMPORTANTE: ID del Google Sheet de agentes
const SHEET_AGENTES_ID = '19Ik0mKdZwN37NZLBV-QLS5JLlBkUHcf3oBiIYr8u8Vo';
const SHEET_AGENTES_NOMBRE = 'RESPUESTAS'; // Nombre de la hoja dentro del archivo

// Hoja Maestra de cursos/cargos
const SHEET_MAESTRA_ID = '1B-TyiGskMJ-TDv2ck4cNyjPFMCJE6p87I3USF2zeFWc';
const SHEET_MAESTRA_NOMBRE = 'Maestra';

const SHEET_SOLICITUDES = 'Solicitudes';
const SHEET_JUSTIFICACIONES = 'Justificaciones';
const FOLDER_ARCHIVOS_ID = '1-48O-hpqqbINKbqdvZPVLT0jjzTCOEBz'; // ID de la carpeta en Drive

// ID de la plantilla de Gmail para emails
const PLANTILLA_SOLICITUD = 'r-6079908409768349520';
const PLANTILLA_JUSTIFICACION = 'r8871403211407668186';

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

        const textoMostrar = `${colC}${colD} ${colE} Sec: ${colA}`.trim();

        cargos.push({
          texto: textoMostrar,
          seccion: colA,
          valorOriginal: { colA, colB, colC, colD, colE, colF }
        });
      }
    }

    // Normalizar datos para evitar problemas de serialización hacia el cliente
    const agentePlano = {
      nombre: agente.nombre || '',
      apellidos: agente.apellidos || '',
      nombres: agente.nombres || '',
      email: agente.email || '',
      dni: agente.dni || '',
      numeroEmpleado: agente.numeroEmpleado || ''
    };

    const cargosPlano = cargos.map(c => ({ texto: c.texto || '', seccion: c.seccion || '' }));

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
      agente.numeroEmpleado,
      agente.apellidos,
      agente.nombres,
      datos.fechaDesde,
      datos.fechaHasta,
      datos.cursoOCargo || '',
      datos.tipoLicencia,
      'Pendiente',
      id
    ]);
    
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
        && data[i][10] === 'Pendiente') {
      solicitudes.push({
        id: data[i][11],
        rowIndex: i + 1,
        tipo: data[i][9],
        desde: Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hasta: Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        cursoOCargo: data[i][8]
      });
    }
  }
  
  return solicitudes;
}

// === GUARDAR JUSTIFICACIÓN ===
function guardarJustificacion(datos, archivoBase64, nombreArchivo, mimeType) {
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
    let folder = DriveApp.getFolderById(FOLDER_ARCHIVOS_ID);

    // Guardar archivo en Drive
    let archivoUrl = "";
    if (archivoBase64 && archivoBase64.includes(",")) {
      const contentType = mimeType || "application/pdf";
      const nuevoNombre = `${agente.dni}-${datos.licenciasIds.join("_")}`; //cambio nombre del archivo

      const blob = Utilities.newBlob(
        Utilities.base64Decode(archivoBase64.split(",")[1]),
        contentType,
        nuevoNombre
      );
      const archivo = folder.createFile(blob);

      // doy permisos.
      archivo.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );

      archivoUrl = archivo.getUrl();
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
      archivoUrl,
    ]);

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