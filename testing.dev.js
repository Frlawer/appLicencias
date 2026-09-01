function testObtenerCargosAgente() {
  const dniONumEmpleado = '4017';
  const agente = obtenerAgente(dniONumEmpleado);
  Logger.log('agente encontrado: ' + (agente ? agente.numeroEmpleado : 'null'));
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

function testObtenerAgente() {
  const dniONumEmpleado = '32282931';
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
