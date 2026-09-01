function doGetv2(e) {
  Logger.log('Endpoint legacy doGetv2 deshabilitado. Parámetros recibidos: ' + JSON.stringify((e && e.parameter) || {}));
  return ContentService.createTextOutput('ENDPOINT_DESHABILITADO');
}