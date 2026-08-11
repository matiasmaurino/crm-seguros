// --- MÓDULO VENCIMIENTOS FP ---
// Se sube un CSV de vencimientos de pólizas de Federación Patronal, y por
// cada fila que todavía no tenga una tarea creada (identificada por N° de
// póliza) se genera una tarea de tipo "Renovacion" con vencimiento 10 días
// antes de la fecha real de vencimiento de la póliza ("Vig. Hasta").
//
// Columnas del CSV (encabezado real, separado por comas):
// Vida Modular, Póliza, Producto, Renovada por, Vig. Desde, Vig. Hasta,
// Matrícula, Asegurado, Contratante, Ramo/Modelo, Saldo, Suma Asegurada,
// Año, Plan, Cód. Productor, Productor

const ID_CARPETA_VENCIMIENTOS_FP = "1abWm1Ue0ZhePajaEroWsubk2qKVvpOfa";

/**
 * Sube el CSV elegido en la web a la carpeta de Drive de Vencimientos FP.
 */
function subirCSVVencimientosFP(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_FP);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, "text/csv", nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoVencimientoFPAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

/**
 * Arma el mapa código de Ramo -> nombre de Ramo, leyendo la hoja AUX del
 * archivo de Liquidaciones (columna A = código, columna B = nombre para
 * Federación Patronal). El código de Ramo de cada fila del CSV de
 * Vencimientos FP es el número que está ANTES de la primera "/" en la
 * columna Póliza (ej: "4/34047237/0" -> código "4").
 *
 * NOTA: usa la constante ID_SPREADSHEET_LIQUIDACIONES, ya definida en
 * Polizas.gs (mismo proyecto) — no hace falta declararla de nuevo acá.
 */
function obtenerMapaRamoFP() {
  const mapa = {};
  try {
    const ssLiq = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
    const hojaAux = ssLiq.getSheetByName("AUX");
    if (!hojaAux) return mapa;

    const datos = hojaAux.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const codigo = (datos[i][0] || "").toString().trim();
      const nombreFP = (datos[i][1] || "").toString().trim(); // Columna B = FP
      if (codigo) mapa[codigo] = nombreFP;
    }
  } catch (e) {
    Logger.log("No se pudo leer la hoja AUX para el mapeo de Ramo: " + e.toString());
  }
  return mapa;
}

/**
 * Registra en la hoja REGISTRO_VENCIMIENTOS_FP (se crea sola si no existe)
 * una fila por cada Responsable que tuvo tareas nuevas en esta corrida,
 * con la cantidad y el rango de fechas de vencimiento real de póliza
 * (Vig. Hasta) que cubrió la tanda importada.
 */
function registrarImportacionVencimientosFP(ss, statsPorResponsable) {
  let hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_FP");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_VENCIMIENTOS_FP");
    hoja.appendRow(["Fecha de Importación", "Responsable", "Cantidad de Tareas", "Vencimiento Más Antiguo", "Vencimiento Más Reciente"]);
  }

  const ahora = new Date();
  Object.keys(statsPorResponsable).forEach(responsable => {
    const stats = statsPorResponsable[responsable];
    if (stats.cantidad === 0) return;
    hoja.appendRow([
      ahora,
      responsable,
      stats.cantidad,
      stats.minFecha,
      stats.maxFecha
    ]);
  });

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila > 1) {
    hoja.getRange(2, 1, ultimaFila - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hoja.getRange(2, 4, ultimaFila - 1, 2).setNumberFormat("dd/mm/yyyy");
  }
}

/**
 * Devuelve el historial completo de importaciones de Vencimientos FP, para
 * mostrarlo en la pantalla (más recientes primero).
 */
function obtenerHistorialVencimientosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_FP");
  if (!hoja) return [];

  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(fila => {
    const fechaImport = (fila[0] instanceof Date) ? Utilities.formatDate(fila[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm") : "";
    const desde = (fila[3] instanceof Date) ? Utilities.formatDate(fila[3], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    const hasta = (fila[4] instanceof Date) ? Utilities.formatDate(fila[4], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    return {
      fechaImportacion: fechaImport,
      responsable: fila[1],
      cantidad: fila[2],
      desde: desde,
      hasta: hasta
    };
  }).reverse();
}

/**
 * Procesa todos los CSV que estén (sin mover a "Procesados" todavía) en la
 * carpeta de Vencimientos FP. Por cada fila cuya póliza no tenga ya una
 * tarea de Renovacion creada, genera una tarea nueva. Al terminar, mueve
 * los archivos ya procesados a "Procesados".
 */
function procesarVencimientosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  const hojaClientes = ss.getSheetByName("CLIENTES");
  const datosClientes = hojaClientes.getDataRange().getValues();
  const clientePorMatricula = {};
  for (let i = 1; i < datosClientes.length; i++) {
    const matricula = (datosClientes[i][8] || "").toString().trim(); // Fed. Patronal (columna I)
    if (matricula) clientePorMatricula[matricula] = datosClientes[i][0]; // id cliente
  }

  // Pólizas que ya tienen tarea de Renovacion creada (columna P de TAREAS)
  const datosTareas = hojaTareas.getDataRange().getValues();
  const polizasYaCreadas = new Set();
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const poliza = (datosTareas[i][15] || "").toString().trim(); // columna P
    if (tipoTarea === "renovacion" && poliza) polizasYaCreadas.add(poliza);
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_FP);
  const iterador = folder.getFilesByType(MimeType.CSV);
  const archivosCSV = [];
  while (iterador.hasNext()) archivosCSV.push(iterador.next());

  if (archivosCSV.length === 0) {
    return { exito: true, tareasCreadas: 0, omitidas: 0, mensaje: "No hay archivos nuevos para procesar." };
  }

  let idTareaActual = 0;
  const ultimaFilaTareas = hojaTareas.getLastRow();
  if (ultimaFilaTareas > 1) {
    idTareaActual = Number(hojaTareas.getRange(ultimaFilaTareas, 2).getValue()) || 0;
  }

  const filasNuevas = [];
  let filasOmitidas = 0;
  const mapaRamoFP = obtenerMapaRamoFP();
  const statsPorResponsable = {
    "Matias": { cantidad: 0, minFecha: null, maxFecha: null },
    "Mauro": { cantidad: 0, minFecha: null, maxFecha: null }
  };

  archivosCSV.forEach(file => {
    const csvData = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8'));

    for (let i = 1; i < csvData.length; i++) {
      const fila = csvData[i];
      if (!fila || fila.length < 16) continue;

      const poliza = (fila[1] || "").toString().trim();          // Póliza
      const producto = (fila[2] || "").toString().trim();        // Producto
      const vigHastaRaw = (fila[5] || "").toString().trim();     // Vig. Hasta
      const matricula = (fila[6] || "").toString().trim();       // Matrícula
      const modeloTexto = (fila[9] || "").toString().trim();     // Ramo/Modelo (en realidad es la descripción del vehículo/bien)
      const productor = (fila[15] || "").toString().trim().toUpperCase(); // Productor

      if (!poliza || polizasYaCreadas.has(poliza)) {
        if (poliza) filasOmitidas++;
        continue;
      }

      // El código de Ramo real es el número antes de la primera "/" en la Póliza
      const codigoRamo = poliza.split("/")[0].trim();
      const ramo = mapaRamoFP[codigoRamo] || "";

      const vigHasta = new Date(vigHastaRaw + "T12:00:00");
      if (isNaN(vigHasta)) continue;

      const fechaVencimientoTarea = new Date(vigHasta);
      fechaVencimientoTarea.setDate(fechaVencimientoTarea.getDate() - 10);

      const idCliente = clientePorMatricula[matricula] || "";
      const responsable = (productor === "MAURIÑO MATIAS") ? "Matias" : "Mauro";

      idTareaActual++;
      filasNuevas.push([
        new Date(),                    // A: Fecha Creación
        idTareaActual,                  // B: ID Tarea
        "Federación Patronal",          // C: Compañía
        "Renovacion",                    // D: Tipo de Tarea
        "Renovación póliza " + poliza + " - " + producto + (modeloTexto ? " - " + modeloTexto : ""), // E: Descripción
        fechaVencimientoTarea,           // F: Vencimiento (10 días antes de Vig. Hasta)
        "Sin leer",                      // G: Estado
        "",                               // H
        "",                               // I
        "Normal",                         // J: Prioridad
        "",                               // K: Adjunto
        idCliente,                        // L: ID Cliente
        "Sistema",                        // M: Usuario
        responsable,                      // N: Responsable
        ramo,                             // O: Ramo (mapeado vía AUX, no el modelo del vehículo)
        poliza                            // P: N° Póliza (mismo campo que "Número Siniestro")
      ]);

      polizasYaCreadas.add(poliza); // por si la misma póliza aparece 2 veces en el mismo CSV

      // Vamos guardando el rango de fechas de vencimiento real (Vig. Hasta)
      // que se importó para cada responsable, para el registro persistente.
      const stats = statsPorResponsable[responsable];
      stats.cantidad++;
      if (!stats.minFecha || vigHasta < stats.minFecha) stats.minFecha = vigHasta;
      if (!stats.maxFecha || vigHasta > stats.maxFecha) stats.maxFecha = vigHasta;
    }

    moverArchivoVencimientoFPAProcesado(file, folder);
  });

  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);

    filasNuevas.forEach((_, idx) => {
      try {
        agendarTareaEnCalendar(filaInicio + idx);
      } catch (e) {
        Logger.log("Error al agendar vencimiento FP en calendario: " + e.toString());
      }
    });

    registrarImportacionVencimientosFP(ss, statsPorResponsable);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas
  };
}

/**
 * Tareas de Renovacion que vienen de Federación Patronal, para el módulo
 * "Vencimientos FP" (se muestra igual que Siniestros).
 */
function obtenerVencimientosFP() {
  const todasLasTareas = obtenerTareas();
  return todasLasTareas.filter(t =>
    (t.tipoTarea || "").toString().trim().toLowerCase() === "renovacion" &&
    (t.compania || "").toString().trim().toLowerCase() === "federación patronal"
  );
}

/**
 * EJECUTAR MANUALMENTE UNA SOLA VEZ: corrige el Ramo de las tareas de
 * Vencimientos FP que ya se crearon con el nombre del vehículo/bien en vez
 * de la categoría real (bug ya corregido en procesarVencimientosFP() para
 * las importaciones nuevas). Recorre TAREAS, y para cada tarea de tipo
 * "Renovacion" y Compañía "Federación Patronal" con N° de Póliza cargado
 * (columna P), recalcula el Ramo correcto vía el código antes de la
 * primera "/" de la Póliza + el mapeo de la hoja AUX, y lo pisa en la
 * columna O si es distinto al que ya tenía.
 */
function corregirRamoVencimientosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) {
    Logger.log("No se encontró la hoja TAREAS.");
    return;
  }

  const mapaRamoFP = obtenerMapaRamoFP();
  if (Object.keys(mapaRamoFP).length === 0) {
    Logger.log("No se pudo armar el mapa de Ramo desde la hoja AUX. No se corrigió nada.");
    return;
  }

  const ultimaFila = hojaTareas.getLastRow();
  if (ultimaFila < 2) return;

  const datos = hojaTareas.getRange(2, 1, ultimaFila - 1, 16).getValues();
  let corregidas = 0;
  let sinCodigoValido = 0;

  for (let i = 0; i < datos.length; i++) {
    const compania = (datos[i][2] || "").toString().trim().toLowerCase();  // Columna C
    const tipoTarea = (datos[i][3] || "").toString().trim().toLowerCase(); // Columna D
    const poliza = (datos[i][15] || "").toString().trim();                  // Columna P

    if (tipoTarea !== "renovacion" || compania !== "federación patronal" || !poliza) continue;

    const codigoRamo = poliza.split("/")[0].trim();
    const ramoCorrecto = mapaRamoFP[codigoRamo] || "";

    if (!ramoCorrecto) {
      sinCodigoValido++;
      continue;
    }

    const ramoActual = (datos[i][14] || "").toString().trim(); // Columna O
    if (ramoCorrecto !== ramoActual) {
      const filaReal = i + 2;
      hojaTareas.getRange(filaReal, 15).setValue(ramoCorrecto); // Columna O
      corregidas++;
    }
  }

  Logger.log(
    "Corrección de Ramo completa. Tareas corregidas: " + corregidas +
    ". Tareas con código de Ramo no encontrado en AUX: " + sinCodigoValido
  );
}