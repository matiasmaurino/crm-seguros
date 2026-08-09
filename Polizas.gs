// --- MÓDULO PÓLIZAS (producción de cada cliente) ---
//
// Combina dos fuentes:
// 1. LIQUIDACIONES AGRUPADAS (archivo externo, oficial) — se actualiza 1 vez
//    por mes, así que una póliza recién emitida puede tardar en aparecer.
// 2. POLIZAS_MANUALES (nuestra propia hoja) — para cargar a mano una póliza
//    que todavía no llegó al archivo oficial (por ejemplo, recién emitida).
//
// Cuando el archivo oficial finalmente trae esa misma póliza (mismo N° +
// compañía), la carga manual se deja de mostrar sola automáticamente — no
// hace falta borrar nada a mano para que no se duplique.
//
// Vinculación con el cliente: la hoja LIQUIDACIONES AGRUPADAS trae, por cada
// movimiento, la compañía (columna CIA) y un número de matrícula (columna
// CLIENTE) — ese número es el mismo que ya guardamos en CLIENTES en las
// columnas Rivadavia / Provincia / Fed. Patronal según la compañía.

const ID_SPREADSHEET_LIQUIDACIONES = "1iocIMPzg31RUv5RxHbtcuTMNt92l4IZUaYFaRwBLqp0";

function nombreCiaLegible(cia) {
  const c = (cia || "").toString().trim().toUpperCase();
  if (c === "FP" || c === "FEDERACION PATRONAL") return "Fed. Patronal";
  if (c === "PS" || c === "PROVINCIA" || c === "PROVINCIA SEGUROS") return "Provincia Seguros";
  if (c === "RIV" || c === "RIVADAVIA") return "Rivadavia";
  return cia || "";
}

// Reduce cualquier variante de escritura de la compañía a un código corto fijo
function normalizarCia(cia) {
  const c = (cia || "").toString().trim().toUpperCase();
  if (c === "FP" || c === "FEDERACION PATRONAL") return "FP";
  if (c === "PS" || c === "PROVINCIA" || c === "PROVINCIA SEGUROS") return "PS";
  if (c === "RIV" || c === "RIVADAVIA") return "RIV";
  return c;
}

// --- POLIZAS_MANUALES: lectura/escritura básica ---

function obtenerPolizasManuales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("POLIZAS_MANUALES");
  if (!hoja) return [];
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map((r, index) => {
    if (!r[2]) return null; // sin cliente, fila vacía
    return {
      id_fila: index + 2,
      fechaCreacion: r[1],
      idCliente: r[2],
      compania: r[3] || "",
      ramo: r[4] || "",
      poliza: (r[5] || "").toString().trim(),
      productor: r[6] || "",
      usuario: r[7] || ""
    };
  }).filter(m => m !== null);
}

function obtenerPolizasManualesPorCliente(idCliente) {
  return obtenerPolizasManuales().filter(m => m.idCliente == idCliente);
}

/**
 * datos = { idCliente, compania, ramo, poliza, productor }
 */
function guardarPolizaManual(datos, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("POLIZAS_MANUALES");
  if (!hoja) throw new Error("No se encontró la hoja POLIZAS_MANUALES");

  const data = hoja.getDataRange().getValues();
  const idNuevo = data.length > 1 ? Number(data[data.length - 1][0]) + 1 : 1;

  hoja.appendRow([
    idNuevo,
    new Date(),
    datos.idCliente,
    datos.compania || "",
    datos.ramo || "",
    (datos.poliza || "").toString().trim(),
    datos.productor || "",
    usuarioActivo || "Sistema"
  ]);

  return { exito: true };
}

function eliminarPolizaManualEnServidor(idFila) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("POLIZAS_MANUALES");
  if (!hoja) throw new Error("No se encontró la hoja POLIZAS_MANUALES");

  const filaNumero = Number(idFila);
  if (filaNumero > 1 && filaNumero <= hoja.getLastRow()) {
    hoja.deleteRow(filaNumero);
    return { exito: true };
  } else {
    throw new Error("Número de fila inválido para eliminar.");
  }
}

/**
 * EJECUTAR MANUALMENTE (por ejemplo, después de actualizar el archivo de
 * Liquidaciones cada mes): borra de POLIZAS_MANUALES las cargas que ya
 * están confirmadas por el archivo oficial (mismo N° de póliza + compañía),
 * para que la hoja no se vaya llenando de registros que ya no hacen falta.
 * No es obligatorio correrla nunca — el sistema ya las oculta solo — pero
 * ayuda a mantener la hoja limpia.
 */
function limpiarPolizasManualesConfirmadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("POLIZAS_MANUALES");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const ssLiq = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
  const hojaLiq = ssLiq.getSheetByName("LIQUIDACIONES AGRUPADAS");
  const datosLiq = hojaLiq ? hojaLiq.getDataRange().getValues() : [];

  const clavesOficiales = new Set();
  for (let i = 1; i < datosLiq.length; i++) {
    const cia = normalizarCia(datosLiq[i][9]);
    const poliza = (datosLiq[i][4] || "").toString().trim();
    if (poliza) clavesOficiales.add(cia + "|" + poliza);
  }

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 8).getValues();
  const filasABorrar = [];

  for (let i = 0; i < datos.length; i++) {
    const cia = normalizarCia(datos[i][3]);
    const poliza = (datos[i][5] || "").toString().trim();
    if (poliza && clavesOficiales.has(cia + "|" + poliza)) {
      filasABorrar.push(i + 2); // número de fila real en la hoja
    }
  }

  // Borramos de abajo hacia arriba para no correr los índices de fila
  filasABorrar.sort((a, b) => b - a).forEach(fila => hoja.deleteRow(fila));

  Logger.log(`Pólizas manuales confirmadas y limpiadas: ${filasABorrar.length}`);
}

// --- CONSULTA POR CLIENTE (pantalla de Clientes) ---

/**
 * Pólizas de UN cliente: cruza sus matrículas contra Liquidaciones oficial,
 * y suma las cargadas a mano que todavía no aparecen ahí.
 */
function obtenerPolizasPorCliente(idCliente) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName("CLIENTES");
  const datosClientes = hojaClientes.getDataRange().getValues();

  let rivadavia = "", provincia = "", fedPatronal = "";
  for (let i = 1; i < datosClientes.length; i++) {
    if (datosClientes[i][0].toString() === idCliente.toString()) {
      rivadavia = (datosClientes[i][6] || "").toString().trim();
      provincia = (datosClientes[i][7] || "").toString().trim();
      fedPatronal = (datosClientes[i][8] || "").toString().trim();
      break;
    }
  }

  const matriculaPorCia = { "FP": fedPatronal, "PS": provincia, "RIV": rivadavia };

  const ssLiq = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
  const hojaLiq = ssLiq.getSheetByName("LIQUIDACIONES AGRUPADAS");
  const datosLiq = hojaLiq ? hojaLiq.getDataRange().getValues() : [];
  // A:PAS B:FECHA C:NOMBRE D:RAMO E:POLIZA F:PRIMA G:PREMIO H:COMISION I:CLIENTE J:CIA K:PAS_AGRUPADO L:RAMO_NOMBRE

  const porPoliza = {};
  const clavesOficiales = new Set();

  for (let i = 1; i < datosLiq.length; i++) {
    const fila = datosLiq[i];
    const ciaCodigo = normalizarCia(fila[9]);
    const poliza = (fila[4] || "").toString().trim();
    if (poliza) clavesOficiales.add(ciaCodigo + "|" + poliza);

    const matriculaEsperada = matriculaPorCia[ciaCodigo];
    const clienteMatricula = (fila[8] || "").toString().trim();
    if (!matriculaEsperada || !clienteMatricula || clienteMatricula !== matriculaEsperada) continue;

    const fechaRaw = fila[1];
    const fechaObj = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);
    const clave = ciaCodigo + "|" + poliza;

    if (!porPoliza[clave] || (fechaObj instanceof Date && !isNaN(fechaObj) && fechaObj > porPoliza[clave]._fechaObj)) {
      porPoliza[clave] = {
        cia: nombreCiaLegible(ciaCodigo),
        ramo: fila[11] || fila[3] || "",
        poliza: poliza,
        prima: fila[5],
        premio: fila[6],
        comision: fila[7],
        pendiente: false,
        _fechaObj: fechaObj
      };
    }
  }

  // Sumamos las cargadas a mano que Liquidaciones todavía no confirmó
  obtenerPolizasManualesPorCliente(idCliente).forEach(m => {
    const ciaCodigo = normalizarCia(m.compania);
    const clave = ciaCodigo + "|" + m.poliza;
    if (!m.poliza || clavesOficiales.has(clave) || porPoliza[clave]) return;

    porPoliza[clave] = {
      cia: nombreCiaLegible(m.compania),
      ramo: m.ramo,
      poliza: m.poliza,
      prima: "",
      premio: "",
      comision: "",
      pendiente: true,
      idFilaManual: m.id_fila,
      _fechaObj: (m.fechaCreacion instanceof Date) ? m.fechaCreacion : new Date()
    };
  });

  const listaPolizas = Object.values(porPoliza);
  listaPolizas.sort((a, b) => b._fechaObj - a._fechaObj);

  return listaPolizas.map(p => {
    const fechaFormat = (p._fechaObj instanceof Date && !isNaN(p._fechaObj))
      ? Utilities.formatDate(p._fechaObj, "GMT-3", "dd/MM/yyyy")
      : "";
    return {
      cia: p.cia,
      ramo: p.ramo,
      poliza: p.poliza,
      ultimaFecha: fechaFormat,
      prima: p.prima,
      premio: p.premio,
      comision: p.comision,
      pendiente: p.pendiente,
      idFilaManual: p.idFilaManual || ""
    };
  });
}

// --- CONSULTA GENERAL (módulo Pólizas) ---

/**
 * Todas las pólizas (Liquidaciones oficial + cargadas a mano no confirmadas
 * todavía), vinculadas a un cliente del CRM cuando se puede.
 */
function obtenerTodasLasPolizas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaClientes = ss.getSheetByName("CLIENTES");
  const datosClientes = hojaClientes.getDataRange().getValues();

  const mapaClientePorMatricula = {};
  const nombrePorIdCliente = {};
  for (let i = 1; i < datosClientes.length; i++) {
    const fila = datosClientes[i];
    const id = fila[0];
    const nombre = fila[1];
    nombrePorIdCliente[id] = nombre;
    const rivadavia = (fila[6] || "").toString().trim();
    const provincia = (fila[7] || "").toString().trim();
    const fedPatronal = (fila[8] || "").toString().trim();
    if (rivadavia) mapaClientePorMatricula["RIV|" + rivadavia] = { id: id, nombre: nombre };
    if (provincia) mapaClientePorMatricula["PS|" + provincia] = { id: id, nombre: nombre };
    if (fedPatronal) mapaClientePorMatricula["FP|" + fedPatronal] = { id: id, nombre: nombre };
  }

  const ssLiq = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
  const hojaLiq = ssLiq.getSheetByName("LIQUIDACIONES AGRUPADAS");
  const datosLiq = hojaLiq ? hojaLiq.getDataRange().getValues() : [];

  const porPoliza = {};
  const clavesOficiales = new Set();

  for (let i = 1; i < datosLiq.length; i++) {
    const fila = datosLiq[i];
    const ciaRaw = fila[9];
    const ciaCodigo = normalizarCia(ciaRaw);
    const matricula = (fila[8] || "").toString().trim();
    const poliza = (fila[4] || "").toString().trim();
    if (!poliza) continue;

    const clave = ciaCodigo + "|" + poliza;
    clavesOficiales.add(clave);

    const fechaRaw = fila[1];
    const fechaObj = (fechaRaw instanceof Date) ? fechaRaw : new Date(fechaRaw);

    if (!porPoliza[clave] || (fechaObj instanceof Date && !isNaN(fechaObj) && fechaObj > porPoliza[clave]._fechaObj)) {
      const clienteVinculado = mapaClientePorMatricula[ciaCodigo + "|" + matricula] || null;
      porPoliza[clave] = {
        productor: fila[0] || "",
        ramo: fila[11] || fila[3] || "",
        poliza: poliza,
        cia: nombreCiaLegible(ciaRaw),
        prima: fila[5],
        premio: fila[6],
        comision: fila[7],
        idCliente: clienteVinculado ? clienteVinculado.id : "",
        clienteNombre: clienteVinculado ? clienteVinculado.nombre : ((fila[2] || "Sin identificar") + " (sin vincular)"),
        pendiente: false,
        _fechaObj: fechaObj
      };
    }
  }

  // Sumamos las cargadas a mano que Liquidaciones todavía no confirmó
  obtenerPolizasManuales().forEach(m => {
    const ciaCodigo = normalizarCia(m.compania);
    const clave = ciaCodigo + "|" + m.poliza;
    if (!m.poliza || clavesOficiales.has(clave) || porPoliza[clave]) return;

    porPoliza[clave] = {
      productor: m.productor || "",
      ramo: m.ramo,
      poliza: m.poliza,
      cia: nombreCiaLegible(m.compania),
      prima: "",
      premio: "",
      comision: "",
      idCliente: m.idCliente,
      clienteNombre: nombrePorIdCliente[m.idCliente] || "Cliente sin nombre",
      pendiente: true,
      idFilaManual: m.id_fila,
      _fechaObj: (m.fechaCreacion instanceof Date) ? m.fechaCreacion : new Date()
    };
  });

  const lista = Object.values(porPoliza);
  lista.sort((a, b) => b._fechaObj - a._fechaObj);

  return lista.map(p => {
    const fechaFormat = (p._fechaObj instanceof Date && !isNaN(p._fechaObj))
      ? Utilities.formatDate(p._fechaObj, "GMT-3", "dd/MM/yyyy")
      : "";
    return {
      productor: p.productor,
      ramo: p.ramo,
      poliza: p.poliza,
      cia: p.cia,
      ultimaFecha: fechaFormat,
      prima: p.prima,
      premio: p.premio,
      comision: p.comision,
      idCliente: p.idCliente,
      clienteNombre: p.clienteNombre,
      pendiente: p.pendiente,
      idFilaManual: p.idFilaManual || ""
    };
  });
}