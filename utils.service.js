/**
 * UTILS SERVICE
 * Funciones de utilidad reutilizables en todo el sistema
 * Incluye: formateo de fechas, conversiones, normalización de datos
 */

/**
 * Formatea una fecha a formato legible (dd/mm/yyyy HH:mm:ss)
 * @param {Date|string} fecha - Fecha a formatear
 * @returns {string} - Fecha formateada o '-' si es inválida
 */
function formatearFechaCompleta_(fecha) {
  const date = fecha instanceof Date ? fecha : new Date(fecha);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '-';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

/**
 * Formatea una fecha a formato corto (dd/mm/yyyy)
 * @param {Date|string} fecha - Fecha a formatear
 * @returns {string} - Fecha formateada o '-' si es inválida
 */
function formatearFechaCorta_(fecha) {
  const date = fecha instanceof Date ? fecha : new Date(fecha);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '-';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * Formatea una fecha a formato ISO (yyyy-MM-dd)
 * @param {Date|string} fecha - Fecha a formatear
 * @returns {string} - Fecha en ISO o '-' si es inválida
 */
function formatearFechaISO_(fecha) {
  const date = fecha instanceof Date ? fecha : new Date(fecha);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return '-';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Convierte un valor a string seguro (sin nulos)
 * @param {any} valor - Valor a convertir
 * @param {string} defecto - Valor por defecto si es nulo
 * @returns {string} - String del valor
 */
function aString_(valor, defecto = '') {
  if (valor === null || valor === undefined) return defecto;
  return String(valor).trim();
}

/**
 * Convierte un valor a número seguro
 * @param {any} valor - Valor a convertir
 * @param {number} defecto - Valor por defecto si es inválido
 * @returns {number} - Número del valor
 */
function aNumero_(valor, defecto = 0) {
  const num = Number(valor);
  return isNaN(num) ? defecto : num;
}

/**
 * Normaliza un texto (trim, lowercase)
 * @param {string} texto - Texto a normalizar
 * @returns {string} - Texto normalizado
 */
function normalizarTexto_(texto) {
  return aString_(texto).toLowerCase().trim();
}

/**
 * Normaliza un DNI (solo dígitos)
 * @param {string} dni - DNI a normalizar
 * @returns {string} - DNI normalizado
 */
function normalizarDNI_(dni) {
  return aString_(dni).replace(/\D/g, '');
}

/**
 * Normaliza un email (lowercase, trim)
 * @param {string} email - Email a normalizar
 * @returns {string} - Email normalizado
 */
function normalizarEmail_(email) {
  return aString_(email).toLowerCase().trim();
}

/**
 * Capitaliza una cadena (primera letra mayúscula)
 * @param {string} texto - Texto a capitalizar
 * @returns {string} - Texto capitalizado
 */
function capitalizar_(texto) {
  const str = aString_(texto);
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

/**
 * Capitaliza un nombre completo (apellido, nombre)
 * @param {string} nombre - Nombre a capitalizar
 * @returns {string} - Nombre capitalizado
 */
function capitalizarNombre_(nombre) {
  return aString_(nombre)
    .split(' ')
    .map(word => capitalizar_(word))
    .join(' ');
}

/**
 * Ordena un array de objetos por fecha descendente
 * @param {Array} array - Array a ordenar
 * @param {string} campos - Campo(s) con fecha(s), separados por coma
 * @returns {Array} - Array ordenado
 */
function ordenarPorFechaDesc_(array, campos = 'timestamp') {
  if (!Array.isArray(array)) return array;
  
  const campo = campos.split(',')[0].trim();
  
  return array.sort((a, b) => {
    const dateA = new Date(a[campo] || 0);
    const dateB = new Date(b[campo] || 0);
    return dateB - dateA;
  });
}

/**
 * Filtra un array por un campo que contenga un valor
 * @param {Array} array - Array a filtrar
 * @param {string} campo - Campo a buscar
 * @param {any} valor - Valor a buscar
 * @param {boolean} exacto - Si es búsqueda exacta o contains
 * @returns {Array} - Array filtrado
 */
function filtrarPor_(array, campo, valor, exacto = false) {
  if (!Array.isArray(array)) return array;
  
  const valorNorm = normalizarTexto_(valor);
  
  return array.filter(item => {
    const itemValor = normalizarTexto_(item[campo]);
    return exacto ? itemValor === valorNorm : itemValor.includes(valorNorm);
  });
}

/**
 * Busca un valor en un array de objetos
 * @param {Array} array - Array a buscar
 * @param {string} campo - Campo a buscar
 * @param {any} valor - Valor exacto a buscar
 * @returns {Object|null} - Objeto encontrado o null
 */
function encontrar_(array, campo, valor) {
  if (!Array.isArray(array)) return null;
  
  return array.find(item => {
    return aString_(item[campo]).toLowerCase() === aString_(valor).toLowerCase();
  }) || null;
}

/**
 * Mapea un array de objetos a un array de un campo específico
 * @param {Array} array - Array a mapear
 * @param {string} campo - Campo a extraer
 * @returns {Array} - Array con solo el campo especificado
 */
function extraerCampo_(array, campo) {
  if (!Array.isArray(array)) return [];
  
  return array
    .map(item => item[campo])
    .filter(v => v !== null && v !== undefined);
}

/**
 * Agrupa un array de objetos por un campo
 * @param {Array} array - Array a agrupar
 * @param {string} campo - Campo para agrupar
 * @returns {Object} - Objeto con grupos
 */
function agruparPor_(array, campo) {
  if (!Array.isArray(array)) return {};
  
  return array.reduce((acc, item) => {
    const key = aString_(item[campo]);
    if (!acc[key]) acc[key] = [];
    acc[key].push(item);
    return acc;
  }, {});
}

/**
 * Cuenta ocurrencias de un valor en un array por campo
 * @param {Array} array - Array a contar
 * @param {string} campo - Campo a contar
 * @param {any} valor - Valor a contar (opcional, si no se pasa cuenta todos)
 * @returns {number} - Cantidad de ocurrencias
 */
function contarPor_(array, campo, valor = null) {
  if (!Array.isArray(array)) return 0;
  
  if (valor === null) {
    return array.length;
  }
  
  return array.filter(item => {
    return aString_(item[campo]).toLowerCase() === aString_(valor).toLowerCase();
  }).length;
}

/**
 * Obtiene valores únicos de un array
 * @param {Array} array - Array a filtrar
 * @returns {Array} - Array con valores únicos
 */
function unicos_(array) {
  if (!Array.isArray(array)) return [];
  return [...new Set(array.map(x => aString_(x)))];
}

/**
 * Pagina un array
 * @param {Array} array - Array a paginar
 * @param {number} pagina - Número de página (1-based)
 * @param {number} tamano - Tamaño de página
 * @returns {Object} - {datos: Array, pagina, tamano, total, totalPaginas}
 */
function paginar_(array, pagina = 1, tamano = 25) {
  if (!Array.isArray(array)) return { datos: [], pagina, tamano, total: 0, totalPaginas: 0 };
  
  const total = array.length;
  const totalPaginas = Math.ceil(total / tamano);
  const paginaValida = Math.max(1, Math.min(pagina, totalPaginas));
  const inicio = (paginaValida - 1) * tamano;
  const fin = inicio + tamano;
  
  return {
    datos: array.slice(inicio, fin),
    pagina: paginaValida,
    tamano,
    total,
    totalPaginas
  };
}

/**
 * Difusión de mensaje a través de log (útil para debugging)
 * @param {string} modulo - Módulo de donde viene el log
 * @param {string} mensaje - Mensaje a loguear
 * @param {any} datos - Datos opcionales a loguear
 */
function logDebug_(modulo, mensaje, datos = null) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  if (datos) {
    Logger.log(`[${timestamp}] ${modulo} → ${mensaje}:`, JSON.stringify(datos));
  } else {
    Logger.log(`[${timestamp}] ${modulo} → ${mensaje}`);
  }
}

/**
 * Obtiene el día de la semana (0-6)
 * @param {Date} fecha - Fecha
 * @returns {number} - 0=domingo, 1=lunes, ...
 */
function obtenerDiaSemana_(fecha) {
  const date = fecha instanceof Date ? fecha : new Date(fecha);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return -1;
  }
  return date.getDay();
}

/**
 * Obtiene la diferencia en días entre dos fechas
 * @param {Date} fecha1 - Primera fecha
 * @param {Date} fecha2 - Segunda fecha
 * @returns {number} - Diferencia en días (positivo si fecha1 > fecha2)
 */
function diferenciaDias_(fecha1, fecha2) {
  const d1 = fecha1 instanceof Date ? fecha1 : new Date(fecha1);
  const d2 = fecha2 instanceof Date ? fecha2 : new Date(fecha2);
  
  if (!(d1 instanceof Date) || isNaN(d1.getTime()) ||
      !(d2 instanceof Date) || isNaN(d2.getTime())) {
    return 0;
  }
  
  return Math.floor((d1 - d2) / (1000 * 60 * 60 * 24));
}

/**
 * Suma días a una fecha
 * @param {Date} fecha - Fecha inicial
 * @param {number} dias - Días a sumar (negativo para restar)
 * @returns {Date} - Nueva fecha
 */
function sumarDias_(fecha, dias) {
  const date = fecha instanceof Date ? new Date(fecha) : new Date(fecha);
  date.setDate(date.getDate() + dias);
  return date;
}
