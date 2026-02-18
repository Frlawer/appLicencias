function getConfigValue_(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value && String(value).trim() ? String(value).trim() : fallback;
}

const SHEET_AGENTES_ID = getConfigValue_('SHEET_AGENTES_ID', '19Ik0mKdZwN37NZLBV-QLS5JLlBkUHcf3oBiIYr8u8Vo');
const SHEET_AGENTES_NOMBRE = getConfigValue_('SHEET_AGENTES_NOMBRE', 'RESPUESTAS');

const SHEET_MAESTRA_ID = getConfigValue_('SHEET_MAESTRA_ID', '1B-TyiGskMJ-TDv2ck4cNyjPFMCJE6p87I3USF2zeFWc');
const SHEET_MAESTRA_NOMBRE = getConfigValue_('SHEET_MAESTRA_NOMBRE', 'Maestra');

const SHEET_SOLICITUDES = getConfigValue_('SHEET_SOLICITUDES', 'SOLICITUDES');
const SHEET_JUSTIFICACIONES = getConfigValue_('SHEET_JUSTIFICACIONES', 'JUSTIFICACIONES');
const SHEET_SISO = getConfigValue_('SHEET_SISO', 'SISO');
const SHEET_NOVEDADES = getConfigValue_('SHEET_NOVEDADES', 'NOVEDADES');
const SHEET_LICENCIAS = getConfigValue_('SHEET_LICENCIAS', 'LICENCIAS');
const FOLDER_ARCHIVOS_ID = getConfigValue_('FOLDER_ARCHIVOS_ID', '1-48O-hpqqbINKbqdvZPVLT0jjzTCOEBz');

const PLANTILLA_SOLICITUD = getConfigValue_('PLANTILLA_SOLICITUD', 'r-6079908409768349520');
const PLANTILLA_JUSTIFICACION = getConfigValue_('PLANTILLA_JUSTIFICACION', 'r8871403211407668186');
const PLANTILLA_ADMIN = getConfigValue_('PLANTILLA_ADMIN', 'r-2767467785340734120');
