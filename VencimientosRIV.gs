// --- MÓDULO VENCIMIENTOS RIV ---
// Se sube el reporte de texto de renovaciones de Rivadavia (formato de
// columnas fijas, NO es un CSV) y por cada renovación que todavía no
// tenga tarea creada (identificada por N° de póliza), se genera una tarea
// de tipo "Renovacion" con vencimiento 10 días antes de la fecha "Hasta".
//
// A diferencia de FP, este reporte NO trae un número de cliente/matrícula
// — solo el nombre del Titular. El cliente se intenta vincular por
// nombre (normalizado); si no hay coincidencia exacta con CLIENTES, la
// tarea se crea igual pero sin cliente vinculado (el nombre queda en la
// descripción para vincularlo a mano si hace falta).
//
// Responsable y Compañía son siempre fijos: "Mauro" y "Rivadavia".
//
// Usa la misma carpeta de Drive que Vencimientos FP, pero solo toma los
// archivos de texto plano (MimeType.PLAIN_TEXT) — los CSV de FP que estén
// en la misma carpeta no los toca.

const ID_CARPETA_VENCIMIENTOS_RIV = "1abWm1Ue0ZhePajaEroWsubk2qKVvpOfa";

// Reconoce la línea de datos de cada renovación dentro del reporte de
// Rivadavia. Ejemplo de línea que matchea:
// "  47-02-849561 26/08/2026    26/08/2027  DI NOTO, MAURO       0221  5912764  ..."
// Grupo 1: Póliza. Grupo 2: fecha "Hasta". Grupo 3: Titular.
// El ancla es el patrón del teléfono (dígitos + exactamente 2 espacios +
// más dígitos), que es lo único con espaciado 100% consistente en el
// reporte — los demás campos varían de un espacio a varios según el largo
// del texto.
const REGEX_RENGLON_VENCIMIENTO_RIV = /^\s*(\S+)\s+\d{2}\/\d{2}\/\d{4}\s+(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+\d{2,4}\s{2}\d{6,8}\s/;

function subirTXTVencimientosRIV(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_RIV);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, "text/plain", nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoVencimientoRIVAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

/**
 * Código de Ramo -> nombre de Ramo para Rivadavia, leyendo la hoja AUX
 * (columna A = código, columna C = nombre para RIV). El código de Ramo de
 * cada renovación es el segmento del medio en la Póliza (ej: "47-02-849561"
 * -> código "02").
 */
function obtenerMapaRamoRIV() {
  const mapa = {};
  try {
    const ssLiq = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
    const hojaAux = ssLiq.getSheetByName("AUX");
    if (!hojaAux) return mapa;

    const datos = hojaAux.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const codigo = (datos[i][0] || "").toString().trim();
      const nombreRIV = (datos[i][2] || "").toString().trim(); // Columna C = RIV
      if (codigo) mapa[codigo] = nombreRIV;
    }
  } catch (e) {
    Logger.log("No se pudo leer AUX para el mapeo de Ramo (RIV): " + e.toString());
  }
  return mapa;
}

/**
 * Normaliza un nombre para poder compararlo: saca etiquetas [RIV/PS/FP],
 * saca comas, colapsa espacios, pasa a mayúsculas.
 */
function normalizarNombreParaMatch(nombre) {
  return nombre.toString()
    .replace(/\s*\[[^\]]*\]/g, "")
    .replace(/,/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function obtenerMapaClientesPorNombre(ss) {
  const hojaClientes = ss.getSheetByName("CLIENTES");
  const mapa = {};
  if (!hojaClientes) return mapa;

  const datos = hojaClientes.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    const id = datos[i][0];
    const nombreNormalizado = normalizarNombreParaMatch(datos[i][1] || "");
    if (nombreNormalizado) mapa[nombreNormalizado] = id;
  }
  return mapa;
}

function registrarImportacionVencimientosRIV(ss, cantidad, minFecha, maxFecha) {
  let hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_RIV");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_VENCIMIENTOS_RIV");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Cantidad de Tareas", "Vencimiento Más Antiguo", "Vencimiento Más Reciente"]);
  }
  if (cantidad > 0) {
    hoja.appendRow([new Date(), cantidad, minFecha, maxFecha]);
    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila, 1, 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hoja.getRange(ultimaFila, 3, 1, 2).setNumberFormat("dd/mm/yyyy");
  }
}

function obtenerHistorialVencimientosRIV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_RIV");
  if (!hoja) return [];

  const data = hoja.getDataRange().getValues();
  if (data.length === 0) return [];

  const primeraEsEncabezado = !(data[0][0] instanceof Date);
  const filas = primeraEsEncabezado ? data.slice(1) : data;

  return filas.map(fila => {
    const fecha = (fila[0] instanceof Date) ? Utilities.formatDate(fila[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm") : "";
    const desde = (fila[2] instanceof Date) ? Utilities.formatDate(fila[2], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    const hasta = (fila[3] instanceof Date) ? Utilities.formatDate(fila[3], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    return { fecha: fecha, cantidad: fila[1] || 0, desde: desde, hasta: hasta };
  }).reverse();
}

function procesarVencimientosRIV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  const mapaRamoRIV = obtenerMapaRamoRIV();
  const mapaClientes = obtenerMapaClientesPorNombre(ss);

  // Pólizas que ya tienen tarea de Renovacion de Rivadavia creada (columna P)
  const datosTareas = hojaTareas.getDataRange().getValues();
  const polizasYaCreadas = new Set();
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const compania = (datosTareas[i][2] || "").toString().trim().toLowerCase();
    const poliza = (datosTareas[i][15] || "").toString().trim();
    if (tipoTarea === "renovacion" && compania === "rivadavia" && poliza) polizasYaCreadas.add(poliza);
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_RIV);
  const iterador = folder.getFilesByType(MimeType.PLAIN_TEXT);
  const archivosTXT = [];
  while (iterador.hasNext()) archivosTXT.push(iterador.next());

  if (archivosTXT.length === 0) {
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

  archivosTXT.forEach(file => {
    const contenido = file.getBlob().getDataAsString('ISO-8859-1');
    const lineas = contenido.split('\n');

    lineas.forEach(linea => {
      const match = linea.match(REGEX_RENGLON_VENCIMIENTO_RIV);
      if (!match) return;

      const poliza = match[1].trim();
      const hastaRaw = match[2].trim();
      const titular = match[3].trim();

      if (!poliza || polizasYaCreadas.has(poliza)) {
        if (poliza) filasOmitidas++;
        return;
      }

      const partesPoliza = poliza.split("-");
      const codigoRamo = partesPoliza.length > 1 ? partesPoliza[1].trim() : "";
      const ramo = mapaRamoRIV[codigoRamo] || "";

      const partesFecha = hastaRaw.split("/"); // dd/mm/yyyy
      const hasta = new Date(partesFecha[2] + "-" + partesFecha[1] + "-" + partesFecha[0] + "T12:00:00");
      if (isNaN(hasta)) return;

      const fechaVencimientoTarea = new Date(hasta);
      fechaVencimientoTarea.setDate(fechaVencimientoTarea.getDate() - 10);

      const nombreNormalizado = normalizarNombreParaMatch(titular);
      const idCliente = mapaClientes[nombreNormalizado] || "";
      if (!idCliente) sinVincular++;

      idTareaActual++;
      filasNuevas.push([
        new Date(),                    // A: Fecha Creación
        idTareaActual,                  // B: ID Tarea
        "Rivadavia",                     // C: Compañía (siempre)
        "Renovacion",                     // D: Tipo de Tarea
        "Renovación póliza " + poliza + (ramo ? " - " + ramo : "") + " - Titular: " + titular, // E: Descripción
        fechaVencimientoTarea,            // F: Vencimiento (10 días antes de "Hasta")
        "Sin leer",                       // G: Estado
        "",                                // H
        "",                                // I
        "Normal",                          // J: Prioridad
        "",                                // K: Adjunto
        idCliente,                         // L: ID Cliente (vacío si no se pudo vincular por nombre)
        "Sistema",                         // M: Usuario
        "Mauro",                           // N: Responsable (siempre)
        ramo,                              // O: Ramo
        poliza                             // P: N° Póliza
      ]);

      polizasYaCreadas.add(poliza); // por si la misma póliza aparece 2 veces en el mismo archivo

      if (!minFecha || hasta < minFecha) minFecha = hasta;
      if (!maxFecha || hasta > maxFecha) maxFecha = hasta;
    });

    moverArchivoVencimientoRIVAProcesado(file, folder);
  });

  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);

    filasNuevas.forEach((_, idx) => {
      try {
        agendarTareaEnCalendar(filaInicio + idx);
      } catch (e) {
        Logger.log("Error al agendar vencimiento RIV en calendario: " + e.toString());
      }
    });

    registrarImportacionVencimientosRIV(ss, filasNuevas.length, minFecha, maxFecha);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas,
    sinVincular: sinVincular
  };
}