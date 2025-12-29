function doGetv2(e) {
  // const ss = SpreadsheetApp.openById("1bJkbRh3oQmr0_O9-HDPvmcPKkaLiQTL1RbFkkXz2eXY");
  const ss = SpreadsheetApp.openById("19Ik0mKdZwN37NZLBV-QLS5JLlBkUHcf3oBiIYr8u8Vo");
  const sh = ss.getSheetByName("RESPUESTAS");

  sh.appendRow([
    new Date(),
    e.parameter.email,
    e.parameter.apellido,
    e.parameter.nombre,
    e.parameter.dni,
    e.parameter.empleado,
    e.parameter.tel
  ]);

  const row = sh.getLastRow();

  // Define la celda en la columna S de esa fila
  const cellS = sh.getRange("S" + row);

  cellS.setFormula(`=UPPER($C${row})&", "&UPPER($D${row})`);

  return ContentService.createTextOutput("OK");
}