// --- MÓDULO VENCIMIENTOS PS (Provincia Seguros) ---
// Se sube el Excel de vencimientos de Provincia Seguros y por cada fila
// cuya póliza todavía no tenga tarea creada (identificada por N° de
// póliza), se genera una tarea de tipo "Renovacion" con vencimiento 10
// días antes de la Fecha Vencimiento real.
//
// Columnas del Excel: Póliza, Ramo, Sucursal, Asegurado, Bien,
// Fecha Vencimiento, Premio. El Ramo ya viene como texto legible (no hace
// falta cruzar códigos como en FP/RIV).
//
// Igual que en RIV: no hay ningún campo de cliente/matrícula, solo el
// nombre del Asegurado — se intenta vincular por nombre (normalizado);
// si no hay coincidencia exacta con CLIENTES, la tarea se crea igual pero
// sin cliente vinculado. Responsable y Compañía son siempre fijos: "Mauro"
// y "Provincia Seguros".
//
// Usa la misma carpeta de Drive que Vencimientos FP y RIV, pero solo toma
// los archivos Excel (.xlsx) — no toca los CSV ni los .txt que estén ahí.
//
// REQUIERE el servicio avanzado "Drive API" habilitado en este proyecto
// de Apps Script (Servicios → Agregar un servicio → Drive API), porque
// para leer un .xlsx hay que convertirlo primero a una Google Sheet
// temporal — Apps Script no puede leer el formato .xlsx directamente.

const ID_CARPETA_VENCIMIENTOS_PS = "1abWm1Ue0ZhePajaEroWsubk2qKVvpOfa";

function subirXLSXVencimientosPS(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_PS);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, MimeType.MICROSOFT_EXCEL, nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoVencimientoPSAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

/**
 * Convierte un archivo .xlsx a una Google Sheet temporal para poder leer
 * sus datos, y borra la copia temporal al terminar.
 */
function leerFilasDeXLSX(file) {
  const resource = { name: "temp_convert_ps", mimeType: MimeType.GOOGLE_SHEETS };
  const tempFile = Drive.Files.create(resource, file.getBlob());
  const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
  const tempSheet = tempSpreadsheet.getSheets()[0];
  const datos = tempSheet.getDataRange().getValues();
  Drive.Files.remove(tempFile.id);
  return datos;
}

function parsearFechaVencimientoPS(valor) {
  if (valor instanceof Date) return valor;
  const partes = valor.toString().trim().split("/");
  if (partes.length !== 3) return null;
  const fecha = new Date(partes[2] + "-" + partes[1] + "-" + partes[0] + "T12:00:00");
  return isNaN(fecha) ? null : fecha;
}

function registrarImportacionVencimientosPS(ss, nombresArchivos, cantidad, minFecha, maxFecha) {
  let hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_PS");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_VENCIMIENTOS_PS");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Archivo(s)", "Cantidad de Tareas", "Vencimiento Más Antiguo", "Vencimiento Más Reciente"]);
  }
  if (cantidad > 0) {
    hoja.appendRow([new Date(), nombresArchivos.join(", "), cantidad, minFecha, maxFecha]);
    const ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila, 1, 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
    hoja.getRange(ultimaFila, 4, 1, 2).setNumberFormat("dd/mm/yyyy");
  }
}

function obtenerHistorialVencimientosPS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_VENCIMIENTOS_PS");
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

function procesarVencimientosPS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  // Reutiliza las funciones ya definidas en VencimientosRIV.gs (mismo proyecto)
  const mapaClientes = obtenerMapaClientesPorNombre(ss);

  const datosTareas = hojaTareas.getDataRange().getValues();
  const polizasYaCreadas = new Set();
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const compania = (datosTareas[i][2] || "").toString().trim().toLowerCase();
    const poliza = (datosTareas[i][15] || "").toString().trim();
    if (tipoTarea === "renovacion" && compania === "provincia seguros" && poliza) polizasYaCreadas.add(poliza);
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_VENCIMIENTOS_PS);
  const iterador = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  const archivosXLSX = [];
  while (iterador.hasNext()) archivosXLSX.push(iterador.next());

  if (archivosXLSX.length === 0) {
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

  archivosXLSX.forEach(file => {
    nombresArchivos.push(file.getName());
    const datos = leerFilasDeXLSX(file);

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      if (!fila || !fila[0]) continue;

      const poliza = fila[0].toString().trim();       // Póliza
      const ramo = (fila[1] || "").toString().trim();  // Ramo (ya viene legible)
      const asegurado = (fila[3] || "").toString().trim(); // Asegurado
      const fechaVencRaw = fila[5];                     // Fecha Vencimiento

      if (!poliza || polizasYaCreadas.has(poliza)) {
        if (poliza) filasOmitidas++;
        continue;
      }

      const fechaVenc = parsearFechaVencimientoPS(fechaVencRaw);
      if (!fechaVenc) continue;

      const fechaVencimientoTarea = new Date(fechaVenc);
      fechaVencimientoTarea.setDate(fechaVencimientoTarea.getDate() - 10);

      const nombreNormalizado = normalizarNombreParaMatch(asegurado);
      const idCliente = mapaClientes[nombreNormalizado] || "";
      if (!idCliente) sinVincular++;

      idTareaActual++;
      filasNuevas.push([
        new Date(),                    // A: Fecha Creación
        idTareaActual,                  // B: ID Tarea
        "Provincia Seguros",             // C: Compañía (siempre)
        "Renovacion",                     // D: Tipo de Tarea
        "Renovación póliza " + poliza + (ramo ? " - " + ramo : "") + " - Asegurado: " + asegurado, // E: Descripción
        fechaVencimientoTarea,            // F: Vencimiento (10 días antes de Fecha Vencimiento)
        "Sin leer",                       // G: Estado
        "",                                // H
        "",                                // I
        "Normal",                          // J: Prioridad
        "",                                // K: Adjunto
        idCliente,                         // L: ID Cliente
        "Sistema",                         // M: Usuario
        "Mauro",                           // N: Responsable (siempre)
        ramo,                              // O: Ramo
        poliza                             // P: N° Póliza
      ]);

      polizasYaCreadas.add(poliza);

      if (!minFecha || fechaVenc < minFecha) minFecha = fechaVenc;
      if (!maxFecha || fechaVenc > maxFecha) maxFecha = fechaVenc;
    }

    moverArchivoVencimientoPSAProcesado(file, folder);
  });

  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);

    filasNuevas.forEach((_, idx) => {
      try {
        agendarTareaEnCalendar(filaInicio + idx);
      } catch (e) {
        Logger.log("Error al agendar vencimiento PS en calendario: " + e.toString());
      }
    });

    registrarImportacionVencimientosPS(ss, nombresArchivos, filasNuevas.length, minFecha, maxFecha);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    omitidas: filasOmitidas,
    sinVincular: sinVincular
  };
}