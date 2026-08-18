// --- MÓDULO SINIESTROS FP ---
// El CSV no trae fila de encabezados. Estructura por columna (0-index):
// 0: (sin uso) 1: (sin uso) 2: Póliza 3: Código de Productor
// 4: Asegurado 5: Fecha Denuncia (sin uso) 6: Fecha Ocurrencia
// 7: Estado 8: Tipo de Siniestro 9: N° de Siniestro 10: Causa
//
// Por cada fila cuyo N° de Siniestro no tenga ya una tarea creada, se
// genera una tarea Tipo="Siniestro" con Vencimiento = Fecha Ocurrencia +
// 11 meses. Compañía siempre "Federación Patronal". Responsable: "Matias"
// si el código de Productor es 21438, si no "Pilar". El Ramo sale del
// código antes de la primera "/" en la Póliza, cruzado contra la columna
// FP de la hoja AUX (mismo mecanismo que Vencimientos FP). El cliente se
// vincula por nombre (Asegurado) contra CLIENTES — si no coincide, la
// tarea se crea igual, sin vincular.

const ID_CARPETA_SINIESTROS_FP = "1uh2TrJnvYRKf056_Z1XjmH3iHkFpcRUU";

function subirCSVSiniestrosFP(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, "text/csv", nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoSiniestrosFPAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

function registrarImportacionSiniestrosFP(ss, nombresArchivos, cantidad, minFecha, maxFecha) {
  let hoja = ss.getSheetByName("REGISTRO_SINIESTROS_FP");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_SINIESTROS_FP");
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

function obtenerHistorialSiniestrosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_SINIESTROS_FP");
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

function procesarSiniestrosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  const mapaRamoFP = obtenerMapaRamoFP(); // ya definida en VencimientosFP.gs
  const mapaClientes = obtenerMapaClientesPorNombre(ss); // ya definida en VencimientosRIV.gs

  // N° de Siniestro que ya tienen tarea creada (columna P de TAREAS)
  const datosTareas = hojaTareas.getDataRange().getValues();
  const siniestrosYaCreados = new Set();
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const compania = (datosTareas[i][2] || "").toString().trim().toLowerCase();
    const numSin = (datosTareas[i][15] || "").toString().trim();
    if (tipoTarea === "siniestro" && compania === "federación patronal" && numSin) siniestrosYaCreados.add(numSin);
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
  const iterador = folder.getFilesByType(MimeType.CSV);
  const archivosCSV = [];
  while (iterador.hasNext()) archivosCSV.push(iterador.next());

  if (archivosCSV.length === 0) {
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

  archivosCSV.forEach(file => {
    nombresArchivos.push(file.getName());
    const csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));

    csvData.forEach(fila => {
      if (!fila || fila.length < 11) return;

      const poliza = (fila[2] || "").toString().trim();
      const productor = (fila[3] || "").toString().trim();
      const asegurado = (fila[4] || "").toString().trim();
      const fechaOcurrenciaRaw = (fila[6] || "").toString().trim(); // formato yyyy-mm-dd
      const tipoSiniestro = (fila[8] || "").toString().trim();
      const numSiniestro = (fila[9] || "").toString().trim();
      const causa = (fila[10] || "").toString().trim();

      if (!numSiniestro || siniestrosYaCreados.has(numSiniestro)) {
        if (numSiniestro) filasOmitidas++;
        return;
      }

      const fechaOcurrencia = new Date(fechaOcurrenciaRaw + "T12:00:00");
      if (isNaN(fechaOcurrencia)) return;

      const codigoRamo = poliza.split("/")[0].trim();
      const ramo = mapaRamoFP[codigoRamo] || "";

      const responsable = (productor === "21438") ? "Matias" : "Pilar";

      const nombreNormalizado = normalizarNombreParaMatch(asegurado);
      const idCliente = mapaClientes[nombreNormalizado] || "";
      if (!idCliente) sinVincular++;

      idTareaActual++;
      filasNuevas.push([
        new Date(),                    // A: Fecha Creación
        idTareaActual,                  // B: ID Tarea
        "Federación Patronal",           // C: Compañía (siempre)
        "Siniestro",                      // D: Tipo de Tarea
        "Siniestro " + numSiniestro + " - Póliza " + poliza +
          (tipoSiniestro ? " - " + tipoSiniestro : "") +
          (causa ? " - " + causa : "") +
          " - Asegurado: " + asegurado,  // E: Descripción
        "",                                // F: Vencimiento — se deja vacío, se carga a mano cuando haga falta
        "Sin leer",                       // G: Estado
        "",                                // H
        "",                                // I
        "Normal",                          // J: Prioridad
        "",                                // K: Adjunto
        idCliente,                         // L: ID Cliente
        "Sistema",                         // M: Usuario
        responsable,                       // N: Responsable (Matias si productor 21438, si no Pilar)
        ramo,                              // O: Ramo
        numSiniestro                       // P: N° Siniestro
      ]);

      siniestrosYaCreados.add(numSiniestro); // por si el mismo siniestro aparece 2 veces en el mismo archivo

      if (!minFecha || fechaOcurrencia < minFecha) minFecha = fechaOcurrencia;
      if (!maxFecha || fechaOcurrencia > maxFecha) maxFecha = fechaOcurrencia;
    });

    moverArchivoSiniestrosFPAProcesado(file, folder);
  });

  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);

    // Ya no agendamos en Calendar acá: sin Vencimiento no hay fecha para
    // el evento, y de paso esto era la parte más lenta del proceso (una
    // llamada a Calendar por cada tarea). Si más adelante alguien le carga
    // un Vencimiento a mano desde la ficha de la tarea, el agendado normal
    // de guardarTarea() se encarga solo en ese momento.
    registrarImportacionSiniestrosFP(ss, nombresArchivos, filasNuevas.length, minFecha, maxFecha);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas,
    sinVincular: sinVincular
  };
}