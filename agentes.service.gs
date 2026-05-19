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

function obtenerAgente(dniONumEmpleado) {
  try {
    const ssExterno = SpreadsheetApp.openById(SHEET_AGENTES_ID);
    const sheet = ssExterno.getSheetByName(SHEET_AGENTES_NOMBRE);

    if (!sheet) {
      throw new Error('No se encontró la hoja "' + SHEET_AGENTES_NOMBRE + '" en el archivo externo');
    }

    const data = sheet.getDataRange().getValues();
    const agentesEncontrados = [];

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

    agentesEncontrados.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return agentesEncontrados[0];
  } catch (error) {
    Logger.log('Error al obtener agente: ' + error.toString());
    throw new Error('Error al acceder a la hoja de agentes: ' + error.toString());
  }
}

function obtenerCargosAgente(dniONumEmpleado) {
  try {
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

    const agentePlano = {
      nombre: String(agente.nombre || ''),
      apellidos: String(agente.apellidos || ''),
      nombres: String(agente.nombres || ''),
      email: String(agente.email || ''),
      dni: String(agente.dni || ''),
      numeroEmpleado: String(agente.numeroEmpleado || ''),
      telefono: String(agente.telefono || '')
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

function obtenerEmailsAgentes(numerosEmpleado) {
  try {
    assertAdminAutorizado_();
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

function obtenerAgentesJornada() {
  try {
    assertAdminAutorizado_();
    const ssMaestra = SpreadsheetApp.openById(SHEET_MAESTRA_ID);
    const sheetPunto = ssMaestra.getSheetByName('.');

    if (!sheetPunto) {
      return { success: false, error: 'No se encontró la hoja "." en la Maestra' };
    }

    const data = sheetPunto.getRange('BH:BI').getValues();
    const agentes = [];

    for (let i = 0; i < data.length; i++) {
      const nombre = data[i][0] ? String(data[i][0]).trim() : '';
      const numeroEmpleado = data[i][1] ? String(data[i][1]).trim() : '';

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

function obtenerCoincidenciasAgentes(consulta, limite) {
  try {
    assertAdminAutorizado_();

    const textoConsulta = normalizarTextoAgente_(consulta);
    const maxResultados = Math.max(1, Math.min(Number(limite) || 12, 50));

    if (!textoConsulta) {
      return { success: true, agentes: [] };
    }

    const ssExterno = SpreadsheetApp.openById(SHEET_AGENTES_ID);
    const sheet = ssExterno.getSheetByName(SHEET_AGENTES_NOMBRE);

    if (!sheet) {
      return { success: false, error: 'No se encontró la hoja "' + SHEET_AGENTES_NOMBRE + '" en el archivo externo' };
    }

    const data = sheet.getDataRange().getValues();
    const coincidenciasNombre = [];
    const coincidenciasOtros = [];

    for (let i = 1; i < data.length; i++) {
      const apellidos = String(data[i][COL_AGENTES.APELLIDOS] || '').trim();
      const nombres = String(data[i][COL_AGENTES.NOMBRES] || '').trim();
      const dni = String(data[i][COL_AGENTES.DNI] || '').trim();
      const numeroEmpleado = String(data[i][COL_AGENTES.NUMERO_EMPLEADO] || '').trim();
      const email = String(data[i][COL_AGENTES.EMAIL] || '').trim();
      const telefono = String(data[i][COL_AGENTES.TELEFONO] || '').trim();
      const telefonoAlt = String(data[i][COL_AGENTES.TELEFONO_ALT] || '').trim();
      const domicilio = String(data[i][COL_AGENTES.DOMICILIO] || '').trim();
      const titulo = String(data[i][COL_AGENTES.TITULO] || '').trim();
      const fechaIngresoRaw = data[i][COL_AGENTES.FECHA_INGRESO_CPEM];
      const fechaIngresoCpem = fechaIngresoRaw instanceof Date
        ? Utilities.formatDate(fechaIngresoRaw, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : (fechaIngresoRaw ? String(fechaIngresoRaw).trim() : '');

      const nombreCompleto = `${apellidos} ${nombres}`.trim();

      const nombreCompletoNorm = normalizarTextoAgente_(nombreCompleto);
      const apellidosNorm = normalizarTextoAgente_(apellidos);
      const nombresNorm = normalizarTextoAgente_(nombres);

      const coincideNombre = nombreCompletoNorm.includes(textoConsulta) || apellidosNorm.includes(textoConsulta) || nombresNorm.includes(textoConsulta);
      if (coincideNombre) {
        coincidenciasNombre.push({
          puntaje: nombreCompletoNorm.startsWith(textoConsulta) ? 0 : 1,
          nombre: nombreCompleto,
          apellidos,
          nombres,
          email,
          dni,
          numeroEmpleado,
          telefono,
          telefonoAlt,
          domicilio,
          titulo,
          fechaIngresoCpem
        });
        if (coincidenciasNombre.length >= maxResultados) {
          break;
        }
        continue;
      }

      const coincideResto = [
        dni,
        numeroEmpleado,
        email,
        telefono,
        telefonoAlt,
        domicilio,
        titulo
      ].map(normalizarTextoAgente_).some(function(campo) {
        return campo.includes(textoConsulta);
      });

      if (coincideResto) {
        coincidenciasOtros.push({
          puntaje: 3,
          nombre: nombreCompleto,
          apellidos,
          nombres,
          email,
          dni,
          numeroEmpleado,
          telefono,
          telefonoAlt,
          domicilio,
          titulo,
          fechaIngresoCpem
        });
      }
    }

    const resultados = coincidenciasNombre.length > 0 ? coincidenciasNombre : coincidenciasOtros;

    resultados.sort(function(a, b) {
      if (a.puntaje !== b.puntaje) return a.puntaje - b.puntaje;
      return a.nombre.localeCompare(b.nombre, 'es');
    });

    return { success: true, agentes: resultados.slice(0, maxResultados) };
  } catch (error) {
    Logger.log('Error al buscar coincidencias de agentes: ' + error.toString());
    return { success: false, error: 'Error al buscar coincidencias de agentes: ' + error.toString(), agentes: [] };
  }
}

function normalizarTextoAgente_(valor) {
  return String(valor || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}
