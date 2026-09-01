/**
 * VALIDATION SERVICE
 * Centraliza todas las funciones de validación del sistema
 * Reutilizable en todos los módulos
 */

/**
 * Valida que una hoja exista en el spreadsheet
 * @param {string} sheetName - Nombre de la hoja a validar
 * @param {Spreadsheet} ss - Spreadsheet (opcional, usa el activo por defecto)
 * @returns {Sheet|null} - La hoja si existe, null si no
 */
function obtenerSheet_(sheetName, ss = null) {
  if (!sheetName || typeof sheetName !== 'string') return null;
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return null;
  return spreadsheet.getSheetByName(sheetName) || null;
}

/**
 * Valida que una hoja exista y lanza error si no
 * @param {string} sheetName - Nombre de la hoja
 * @throws {Error} Si la hoja no existe
 */
function assertSheetExists_(sheetName) {
  const sheet = obtenerSheet_(sheetName);
  if (!sheet) {
    throw new Error(`Hoja requerida no encontrada: ${sheetName}`);
  }
  return sheet;
}

/**
 * Valida que un array no esté vacío
 * @param {Array} array - Array a validar
 * @param {string} nombreArray - Nombre del array (para mensaje de error)
 * @returns {boolean} - true si es válido
 * @throws {Error} Si el array está vacío o no es array
 */
function assertArrayNotEmpty_(array, nombreArray = 'Array') {
  if (!Array.isArray(array) || array.length === 0) {
    throw new Error(`${nombreArray} vacío o inválido`);
  }
  return true;
}

/**
 * Valida que un valor no sea nulo o vacío
 * @param {any} valor - Valor a validar
 * @param {string} nombreCampo - Nombre del campo (para mensaje de error)
 * @returns {boolean} - true si es válido
 * @throws {Error} Si el valor es nulo o vacío
 */
function assertNotEmpty_(valor, nombreCampo = 'Campo') {
  if (valor === null || valor === undefined || String(valor).trim() === '') {
    throw new Error(`${nombreCampo} es requerido`);
  }
  return true;
}

/**
 * Valida un timestamp
 * @param {any} timestamp - Timestamp a validar
 * @returns {Date|null} - Date si es válido, null si no
 */
function validarTimestamp_(timestamp) {
  if (timestamp instanceof Date && !isNaN(timestamp.getTime())) {
    return timestamp;
  }
  
  const parsed = new Date(timestamp);
  if (parsed instanceof Date && !isNaN(parsed.getTime())) {
    return parsed;
  }
  
  return null;
}

/**
 * Valida un email
 * @param {string} email - Email a validar
 * @returns {boolean} - true si es válido
 */
function esEmailValido_(email) {
  if (!email || typeof email !== 'string') return false;
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email.trim());
}

/**
 * Valida que un índice de fila sea válido para una sheet
 * @param {number} rowIndex - Índice de fila (1-based)
 * @param {Sheet} sheet - Sheet a validar contra
 * @returns {boolean} - true si es válido
 * @throws {Error} Si rowIndex es inválido
 */
function assertRowIndexValido_(rowIndex, sheet) {
  if (!rowIndex || typeof rowIndex !== 'number') {
    throw new Error('Índice de fila inválido');
  }
  
  if (rowIndex < 2) {
    throw new Error('No se puede modificar el encabezado (fila 1)');
  }
  
  if (rowIndex > sheet.getLastRow()) {
    throw new Error(`Fila ${rowIndex} no existe`);
  }
  
  return true;
}

/**
 * Valida que un DNI tenga formato válido
 * @param {string} dni - DNI a validar
 * @returns {boolean} - true si es válido
 */
function esDNIValido_(dni) {
  if (!dni || typeof dni !== 'string') return false;
  const dniLimpio = dni.trim().replace(/\D/g, '');
  return dniLimpio.length === 8;
}

/**
 * Valida un estado de solicitud contra valores permitidos
 * @param {string} estado - Estado a validar
 * @param {Array} estadosPermitidos - Estados permitidos (opcional)
 * @returns {boolean} - true si es válido
 */
function esEstadoValido_(estado, estadosPermitidos = null) {
  if (!estado || typeof estado !== 'string') return false;
  
  // Estados por defecto permitidos
  const estadosPorDefecto = [
    'Pendiente',
    'Autorizada',
    'Autorizado',
    'Justificada',
    'Completada',
    'Injustificada',
    'Descuento',
    'Rechazada'
  ];
  
  const permitidos = estadosPermitidos || estadosPorDefecto;
  const estadoNormalizado = estado.trim();
  
  return permitidos.some(e => e.toLowerCase() === estadoNormalizado.toLowerCase());
}

/**
 * Valida que un rango de fechas sea válido
 * @param {Date} fechaDesde - Fecha inicial
 * @param {Date} fechaHasta - Fecha final
 * @returns {boolean} - true si es válido
 * @throws {Error} Si las fechas son inválidas
 */
function assertFechasValidas_(fechaDesde, fechaHasta) {
  const desde = validarTimestamp_(fechaDesde);
  const hasta = validarTimestamp_(fechaHasta);
  
  if (!desde || !hasta) {
    throw new Error('Las fechas deben ser válidas');
  }
  
  if (desde > hasta) {
    throw new Error('La fecha desde no puede ser posterior a la fecha hasta');
  }
  
  return true;
}

/**
 * Valida que un número esté en un rango permitido
 * @param {number} numero - Número a validar
 * @param {number} minimo - Valor mínimo (inclusive)
 * @param {number} maximo - Valor máximo (inclusive)
 * @param {string} nombreCampo - Nombre del campo
 * @returns {boolean} - true si es válido
 * @throws {Error} Si está fuera de rango
 */
function assertEnRango_(numero, minimo, maximo, nombreCampo = 'Valor') {
  const num = Number(numero);
  if (isNaN(num) || num < minimo || num > maximo) {
    throw new Error(`${nombreCampo} debe estar entre ${minimo} y ${maximo}`);
  }
  return true;
}

/**
 * Valida un tipo de licencia
 * @param {string} tipoLicencia - Tipo de licencia a validar
 * @returns {boolean} - true si es válido
 */
function esTipoLicenciaValido_(tipoLicencia) {
  if (!tipoLicencia || typeof tipoLicencia !== 'string') return false;
  
  const tiposPermitidos = [
    'Ordinaria',
    'Matrimonio',
    'Familiar',
    'Amamantamiento',
    'Enfermedad',
    'Estudio',
    'Union Estable',
    'Jornada Institucional'
  ];
  
  return tiposPermitidos.some(
    t => t.toLowerCase() === tipoLicencia.trim().toLowerCase()
  );
}

/**
 * Valida que el usuario sea admin autorizado
 * @returns {boolean} - true si es autorizado
 * @throws {Error} Si no tiene permisos
 */
function assertAdminAutorizado() {
  if (!esAdminAutorizado_()) {
    throw new Error('No tenés permisos suficientes para esta operación');
  }
  return true;
}
