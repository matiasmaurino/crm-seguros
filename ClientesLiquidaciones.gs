function sincronizarYLimpiarCRM() {
  const ssCRM = SpreadsheetApp.getActiveSpreadsheet();
  const hojaCRM = ssCRM.getSheetByName("CLIENTES"); 
  
  const idLiquidaciones = "1iocIMPzg31RUv5RxHbtcuTMNt92l4IZUaYFaRwBLqp0";
  const ssLiq = SpreadsheetApp.openById(idLiquidaciones);
  const hojaAgrupados = ssLiq.getSheetByName("CLIENTES AGRUPADOS");
  
  // 1. Obtener datos de referencia de Liquidaciones (ESTA ES LA FUENTE PRINCIPAL)
  const datosLiq = hojaAgrupados.getDataRange().getValues();
  let mapaReferencia = {};
  for (let i = 1; i < datosLiq.length; i++) {
    let nombre = datosLiq[i][1] ? datosLiq[i][1].toString().trim().toUpperCase() : "";
    if (nombre) { 
      mapaReferencia[nombre] = datosLiq[i];
    }
  }
  
  // 2. Obtener TODO el CRM actual para no perder datos manuales
  const rangoCompleto = hojaCRM.getDataRange();
  const datosCRM = rangoCompleto.getValues();
  let clientesProcesados = {}; 

  // 3. Primero, mapeamos lo que ya existe en el CRM
  for (let j = 1; j < datosCRM.length; j++) {
    let colA = datosCRM[j][0] ? datosCRM[j][0].toString().trim().toUpperCase() : "";
    let colB = datosCRM[j][1] ? datosCRM[j][1].toString().trim().toUpperCase() : "";
    let colC = datosCRM[j][2] ? datosCRM[j][2].toString().trim().toUpperCase() : "";
    let nombreReal = mapaReferencia[colB] ? colB : (mapaReferencia[colC] ? colC : (mapaReferencia[colA] ? colA : ""));
    if (nombreReal) {
      let ref = mapaReferencia[nombreReal];
      clientesProcesados[nombreReal] = [
        ref[0], // A: ID
        ref[1], // B: Nombre
        ref[8], // C: CUIL
        ref[5], // D: Domicilio
        ref[6], // E: Telefono
        ref[7], // F: Email
        ref[2] || datosCRM[j][6] || "", // G: Rivadavia
        ref[3] || datosCRM[j][7] || "", // H: Provincia
        ref[4] || datosCRM[j][8] || ""  // I: FedPatronal
      ];
    }
  }

  // 4. Agregar clientes de Liquidaciones que no estaban en absoluto en el CRM
  for (let nombreRef in mapaReferencia) {
    if (!clientesProcesados[nombreRef]) {
      let d = mapaReferencia[nombreRef];
      clientesProcesados[nombreRef] = [
        d[0], d[1], d[8], d[5], d[6], d[7], d[2], d[3], d[4]
      ];
    }
  }
  
  // 5. Preparar resultado final ordenado por nombre
  let resultadoFinal = Object.values(clientesProcesados);
  resultadoFinal.sort((a, b) => a[1].localeCompare(b[1]));
  
  // 6. Limpieza y escritura final
  if (hojaCRM.getLastRow() > 1) {
    hojaCRM.getRange(2, 1, hojaCRM.getLastRow() - 1, 9).clearContent();
  }
  
  if (resultadoFinal.length > 0) {
    hojaCRM.getRange(2, 1, resultadoFinal.length, 9).setValues(resultadoFinal);
  }
  
  ssCRM.toast("Sincronización completa: IDs y números de cliente actualizados.", "Éxito");
}

