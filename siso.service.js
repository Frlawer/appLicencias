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

function obtenerAgentesSisoFiltrados() {
  try {
    assertAdminAutorizado_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SISO);

    if (!sheet) {
      return { success: true, agentes: [], total: 0, mensaje: 'No existe la hoja SISO' };
    }

    const ultimaFila = sheet.getLastRow();
    const ultimaColumna = Math.max(sheet.getLastColumn(), 12);

    if (ultimaFila < 2) {
      return { success: true, agentes: [], total: 0, mensaje: 'La hoja SISO no tiene datos' };
    }

    const valores = sheet.getRange(1, 1, ultimaFila, ultimaColumna).getValues();
    const encabezados = valores[0] || [];
    const filas = valores.slice(1);
    const indiceAgente = obtenerIndiceAgenteSiso_(encabezados);

    const hoy = new Date();
    const inicio = limpiarHoraFecha_(sumarDias_(hoy, -7));
    const fin = limpiarHoraFecha_(sumarDias_(hoy, 30));

    const conceptosPermitidos = {
      '4b - largo tratam.docente (cod.rrhh: 2404b)': true,
      '62 - largo trat. administrativo (cod.rrhh: 2262)': true
    };

    const agentes = filas.reduce((acumulado, fila, index) => {
      const valorI = fila[8];
      const valorK = normalizarComparacionSiso_(fila[10]);
      const valorL = normalizarComparacionSiso_(fila[11]);

      const estadoValido = valorK === 'justificado' || valorK === 'justificacion';
      const conceptoValido = !!conceptosPermitidos[valorL];

      if (!estadoValido || !conceptoValido) {
        return acumulado;
      }

      const fecha = parsearFechaSiso_(valorI);
      if (!fecha) {
        return acumulado;
      }

      const fechaSinHora = limpiarHoraFecha_(fecha);
      if (fechaSinHora < inicio || fechaSinHora > fin) {
        return acumulado;
      }

      acumulado.push({
        agente: obtenerNombreAgenteSiso_(fila, indiceAgente),
        nEmpleado: aString_(fila[1]),
        fecha: Utilities.formatDate(fechaSinHora, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        fechaTexto: Utilities.formatDate(fechaSinHora, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        conceptoL: aString_(fila[11]),
        filaHoja: index + 2
      });

      return acumulado;
    }, []);

    agentes.sort((a, b) => a.fecha.localeCompare(b.fecha) || a.agente.localeCompare(b.agente));

    return {
      success: true,
      agentes,
      total: agentes.length,
      rango: {
        desde: Utilities.formatDate(inicio, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        hasta: Utilities.formatDate(fin, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      }
    };
  } catch (error) {
    Logger.log('Error al obtener agentes filtrados SISO: ' + error.toString());
    return { success: false, error: error.toString(), agentes: [], total: 0 };
  }
}

function obtenerIndiceAgenteSiso_(encabezados) {
  if (!Array.isArray(encabezados) || !encabezados.length) {
    return 1;
  }

  const posibles = ['agente', 'docente', 'apellido', 'nombre'];
  const indice = encabezados.findIndex(celda => {
    const valor = normalizarTexto_(celda);
    return posibles.some(tag => valor.includes(tag));
  });

  return indice >= 0 ? indice : 1;
}

function obtenerNombreAgenteSiso_(fila, indiceAgente) {
  if (!Array.isArray(fila)) return '-';

  const nombrePrincipal = aString_(fila[indiceAgente]);
  if (nombrePrincipal) return nombrePrincipal;

  const respaldo = aString_(fila[1]) || aString_(fila[0]) || aString_(fila[2]);
  return respaldo || '-';
}

function parsearFechaSiso_(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }

  if (typeof valor === 'number' && !isNaN(valor)) {
    const fechaBase = new Date(Date.UTC(1899, 11, 30));
    fechaBase.setUTCDate(fechaBase.getUTCDate() + Math.floor(valor));
    return new Date(fechaBase.getUTCFullYear(), fechaBase.getUTCMonth(), fechaBase.getUTCDate());
  }

  const texto = aString_(valor);
  if (!texto) {
    return null;
  }

  const formatoLatino = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (formatoLatino) {
    const dia = Number(formatoLatino[1]);
    const mes = Number(formatoLatino[2]) - 1;
    const anio = Number(formatoLatino[3]);
    const fechaLatina = new Date(anio, mes, dia);
    if (!isNaN(fechaLatina.getTime())) {
      return fechaLatina;
    }
  }

  const parseada = new Date(texto);
  if (parseada instanceof Date && !isNaN(parseada.getTime())) {
    return parseada;
  }

  return null;
}

function limpiarHoraFecha_(fecha) {
  return new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
}

function normalizarComparacionSiso_(valor) {
  return aString_(valor)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}
