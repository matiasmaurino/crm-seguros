// --- MÓDULO SINIESTROS PS (Provincia Seguros) ---
// El archivo es .xls (formato viejo de Excel), con una fila de título
// arriba de todo ("F. Venta - 10038:MAURIÑO..."), después la fila de
// encabezados reales, y recién ahí los datos:
// Ramo, Siniestro, Nro. Poliza Certificado, Cliente, Patente, Vig. Desde,
// Vig. Hasta, Causa, F.Ingreso, F.Declaracion, F.Ocurrencia
//
// El Ramo ya viene como texto legible (no hace falta cruzar con AUX).
// Compañía siempre "Provincia Seguros", Responsable siempre "Pilar".
// Vencimiento queda vacío (se carga a mano cuando haga falta, igual que
// en Siniestros FP). El cliente se vincula por nombre (Cliente) contra
// CLIENTES — si no coincide, la tarea se crea igual, sin vincular.
//
// Comparte carpeta de Drive con Siniestros FP, pero se filtra por
// nombre de archivo terminado en ".xls" (el tipo MIME de este formato
// viejo es menos confiable para filtrar que la extensión).

/**
 * Convierte un archivo .xls/.xlsx a una Google Sheet temporal para poder
 * leer sus datos, y borra la copia temporal al terminar. Ya existía en
 * VencimientosPS.gs con este mismo nombre — se reutiliza tal cual.
 */
// function leerFilasDeXLSX(file) { ... } -> definida en VencimientosPS.gs

function subirXLSSiniestrosPS(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, "application/vnd.ms-excel", nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoSiniestrosPSAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

function parsearFechaDDMMYYYYSiniestrosPS(valor) {
  if (valor instanceof Date) return valor;
  if (!valor) return null;
  const partes = valor.toString().trim().split("/");
  if (partes.length !== 3) return null;
  const fecha = new Date(partes[2] + "-" + partes[1] + "-" + partes[0] + "T12:00:00");
  return isNaN(fecha) ? null : fecha;
}

function registrarImportacionSiniestrosPS(ss, nombresArchivos, cantidad, minFecha, maxFecha) {
  let hoja = ss.getSheetByName("REGISTRO_SINIESTROS_PS");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_SINIESTROS_PS");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Archivo(s)", "Cantidad de Tareas", "Ocurrencia Más Antigua", "Ocurrencia Más Reciente"]);
  }
  if (cantidad > 0) {
    hoja.appendRow([new Date(), nombresArchivos.join(", "), cantidad, minFecha, maxFecha]);
    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila, 1, 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hoja.getRange(ultimaFila, 4, 1, 2).setNumberFormat("dd/mm/yyyy");
  }
}

function obtenerHistorialSiniestrosPS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_SINIESTROS_PS");
  if (!hoja) return [];

  const data = hoja.getDataRange().getValues();
  if (data.length === 0) return [];

  const primeraEsEncabezado = !(data[0][0] instanceof Date);
  const filas = primeraEsEncabezado ? data.slice(1) : data;

  return filas.map(fila => {
    const fecha = (fila[0] instanceof Date) ? Utilities.formatDate(fila[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm") : "";
    const desde = (fila[3] instanceof Date) ? Utilities.formatDate(fila[3], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    const hasta = (fila[4] instanceof Date) ? Utilities.formatDate(fila[4], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    return { fecha: fecha, archivos: fila[1] || "", cantidad: fila[2] || 0, desde: desde, hasta: hasta };
  }).reverse();
}

function procesarSiniestrosPS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  const mapaClientes = obtenerMapaClientesPorNombre(ss); // ya definida en VencimientosRIV.gs

  // N° de Siniestro que ya tienen tarea creada (columna P de TAREAS)
  const datosTareas = hojaTareas.getDataRange().getValues();
  const siniestrosYaCreados = new Set();
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const compania = (datosTareas[i][2] || "").toString().trim().toLowerCase();
    const numSin = (datosTareas[i][15] || "").toString().trim();
    if (tipoTarea === "siniestro" && compania === "provincia seguros" && numSin) siniestrosYaCreados.add(numSin);
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
  const iterador = folder.getFiles();
  const archivosXLS = [];
  while (iterador.hasNext()) {
    const f = iterador.next();
    if (f.getName().toLowerCase().endsWith('.xls') || f.getName().toLowerCase().endsWith('.xlsx')) {
      archivosXLS.push(f);
    }
  }

  if (archivosXLS.length === 0) {
    return { exito: true, tareasCreadas: 0, omitidas: 0, sinVincular: 0, mensaje: "No hay archivos nuevos para procesar." };
  }

  let idTareaActual = 0;
  const ultimaFilaTareas = hojaTareas.getLastRow();
  if (ultimaFilaTareas > 1) {
    idTareaActual = Number(hojaTareas.getRange(ultimaFilaTareas, 2).getValue()) || 0;
  }

  const filasNuevas = [];
  let filasOmitidas = 0;
  let sinVincular = 0;
  let minFecha = null;
  let maxFecha = null;
  const nombresArchivos = [];

  archivosXLS.forEach(file => {
    nombresArchivos.push(file.getName());
    const datos = leerFilasDeXLSX(file);

    datos.forEach(fila => {
      if (!fila || fila.length < 11) return;

      const numSiniestroRaw = fila[1];
      if (numSiniestroRaw === "" || numSiniestroRaw === null || numSiniestroRaw === undefined) return;
      if (numSiniestroRaw.toString().trim().toLowerCase() === "siniestro") return; // fila de encabezado

      const ramo = (fila[0] || "").toString().trim();
      const numSiniestro = (typeof numSiniestroRaw === "number") ? Math.round(numSiniestroRaw).toString() : numSiniestroRaw.toString().trim();
      const poliza = (fila[2] || "").toString().trim();
      const asegurado = (fila[3] || "").toString().replace(/,\s*$/, "").trim(); // saca la coma al final
      const patente = (fila[4] || "").toString().trim();
      const causa = (fila[7] || "").toString().trim();
      const fechaOcurrenciaRaw = fila[10];

      if (!numSiniestro || siniestrosYaCreados.has(numSiniestro)) {
        filasOmitidas++;
        return;
      }

      const fechaOcurrencia = parsearFechaDDMMYYYYSiniestrosPS(fechaOcurrenciaRaw);

      const responsable = "Pilar";

      const nombreNormalizado = normalizarNombreParaMatch(asegurado);
      const idCliente = mapaClientes[nombreNormalizado] || "";
      if (!idCliente) sinVincular++;

      idTareaActual++;
      filasNuevas.push([
        new Date(),                    // A: Fecha Creación
        idTareaActual,                  // B: ID Tarea
        "Provincia Seguros",             // C: Compañía (siempre)
        "Siniestro",                      // D: Tipo de Tarea
        "Siniestro " + numSiniestro + " - Póliza " + poliza +
          (patente ? " - Patente " + patente : "") +
          (causa && causa.toUpperCase() !== "NO APLICA" ? " - " + causa : "") +
          " - Asegurado: " + asegurado,  // E: Descripción
        "",                                // F: Vencimiento — vacío, se carga a mano
        "Sin leer",                       // G: Estado
        "",                                // H
        "",                                // I
        "Normal",                          // J: Prioridad
        "",                                // K: Adjunto
        idCliente,                         // L: ID Cliente
        "Sistema",                         // M: Usuario
        responsable,                       // N: Responsable (siempre Pilar)
        ramo,                              // O: Ramo
        numSiniestro                       // P: N° Siniestro
      ]);

      siniestrosYaCreados.add(numSiniestro);

      if (fechaOcurrencia) {
        if (!minFecha || fechaOcurrencia < minFecha) minFecha = fechaOcurrencia;
        if (!maxFecha || fechaOcurrencia > maxFecha) maxFecha = fechaOcurrencia;
      }
    });

    moverArchivoSiniestrosPSAProcesado(file, folder);
  });

  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);

    registrarImportacionSiniestrosPS(ss, nombresArchivos, filasNuevas.length, minFecha, maxFecha);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas,
    sinVincular: sinVincular
  };
}