function enviarEmailSolicitud(agente, datos, timestamp, idSolicitud) {
  try {
    const asunto = `Solicitud de Licencia ${agente.nombre} - CPEM N° 25`;
    const draftId = PLANTILLA_SOLICITUD;

    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);

    const fechaCargaFormateada = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    cuerpoHTML = cuerpoHTML.replace('%NOMBRE%', agente.nombre).replace('%TIPO%', datos.tipoLicencia).replace('%FECHA_CARGA%', fechaCargaFormateada);

    GmailApp.createDraft(agente.email, asunto, '', {
      htmlBody: cuerpoHTML
    });

    Logger.log('Borrador de email creado exitosamente para: ' + agente.email);
  } catch (error) {
    Logger.log('Error al crear borrador de solicitud: ' + error.toString());
    throw error;
  }
}

function enviarEmailJustificacion(agente, cantidadLicencias, archivoUrl, timestamp) {
  try {
    const asunto = `Presentación Constancia ${agente.nombre} - CPEM N° 25`;
    const draftId = PLANTILLA_JUSTIFICACION;

    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);

    const fechaCargaFormateada = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    cuerpoHTML = cuerpoHTML
      .replace('%NOMBRE%', agente.nombre)
      .replace('%CANTIDAD_LICENCIAS%', cantidadLicencias)
      .replace('%FECHA_CARGA%', fechaCargaFormateada)
      .replace('%ARCHIVO_URL%', archivoUrl);

    GmailApp.createDraft(agente.email, asunto, '', {
      attachments: [],
      htmlBody: cuerpoHTML
    });

    Logger.log('Borrador de email creado exitosamente para: ' + agente.email);
  } catch (error) {
    Logger.log('Error al crear borrador de justificación: ' + error.toString());
    throw error;
  }
}

function enviarEmailAdmin(emailDocente, nombreDocente, mensaje) {
  try {
    assertAdminAutorizado_();
    const asunto = `Respuesta Licencia ${nombreDocente} - CPEM N° 25`;
    const draftId = PLANTILLA_ADMIN;

    const borrador = GmailApp.getDraft(draftId);
    const cuerpoBase = borrador.getMessage().getBody();

    let cuerpoHTML = String(cuerpoBase);

    cuerpoHTML = cuerpoHTML.replace(/%RESPUESTA%/g, mensaje).replace(/%NOMBRE%/g, nombreDocente);

    GmailApp.createDraft(emailDocente, asunto, '', {
      htmlBody: cuerpoHTML
    });

    return { success: true };
  } catch (error) {
    Logger.log('Error al crear borrador de email: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
