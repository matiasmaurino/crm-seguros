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
function registrarImportacionVencimientosFP(ss, statsPorProductor) {
  let hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_FP");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_VENCIMIENTOS_FP");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Productor", "Cantidad de Tareas", "Vencimiento Más Antiguo", "Vencimiento Más Reciente"]);
  }

  const ahora = new Date();
  Object.keys(statsPorProductor).forEach(productor => {
    const stats = statsPorProductor[productor];
    if (stats.cantidad === 0) return;
    hoja.appendRow([
      ahora,
      productor,
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
  if (data.length === 0) return [];

  // No asumimos por posición que la fila 1 es encabezado: lo detectamos
  // mirando si la primera celda es una fecha real (dato) o texto (título).
  const primeraEsEncabezado = data[0][0] === "" || !(data[0][0] instanceof Date);
  const filas = primeraEsEncabezado ? data.slice(1) : data;

  return filas.map(fila => {
    const fechaImport = (fila[0] instanceof Date) ? Utilities.formatDate(fila[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm") : "";
    const desde = (fila[3] instanceof Date) ? Utilities.formatDate(fila[3], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    const hasta = (fila[4] instanceof Date) ? Utilities.formatDate(fila[4], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    return {
      fechaImportacion: fechaImport,
      productor: fila[1],
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
  let filasSufijoInvalido = 0;
  const mapaRamoFP = obtenerMapaRamoFP();
  const statsPorProductor = {}; // se arma solo, una clave por cada Productor real que aparezca en el CSV

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

      if (!poliza) continue;

      // Solo importamos el movimiento base de la póliza (termina en "/0").
      // Los que terminan en "/1", "/2", "/10", etc. son endosos o
      // movimientos posteriores de la misma póliza, no se cargan.
      const partesPoliza = poliza.split("/");
      const sufijoMovimiento = partesPoliza[partesPoliza.length - 1].trim();
      if (sufijoMovimiento !== "0") {
        filasSufijoInvalido++;
        continue;
      }

      if (polizasYaCreadas.has(poliza)) {
        filasOmitidas++;
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
      // que se importó para cada Productor real (tal cual viene en el CSV),
      // para el registro persistente — cada archivo suele traer un solo
      // Productor, así que esto queda más trazable que agrupar por
      // Responsable interno (Matias/Mauro).
      if (!statsPorProductor[productor]) {
        statsPorProductor[productor] = { cantidad: 0, minFecha: null, maxFecha: null };
      }
      const stats = statsPorProductor[productor];
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

    registrarImportacionVencimientosFP(ss, statsPorProductor);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas,
    noBase: filasSufijoInvalido
  };
}

/**
 * Tareas de Renovacion que vienen de Federación Patronal, para el módulo
 * "Vencimientos FP" (se muestra igual que Siniestros).
 */
// Nota: a pesar del nombre (que quedó de cuando esto era solo FP), ahora
// devuelve renovaciones de CUALQUIER compañía (FP, RIV, etc.) — la
// pantalla "Vencimientos" ya no es exclusiva de Federación Patronal. No
// renombré la función para no romper la referencia que usa el frontend.
function obtenerVencimientosFP() {
  const todasLasTareas = obtenerTareas();
  return todasLasTareas.filter(t =>
    (t.tipoTarea || "").toString().trim().toLowerCase() === "renovacion" &&
    (t.estado || "").toString().trim().toLowerCase() !== "terminado"
  );
}