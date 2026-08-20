// --- MÓDULO SINIESTROS FP ---
// El CSV no trae fila de encabezados. Estructura por columna (0-index):
// 0: (sin uso) 1: (sin uso) 2: Póliza 3: Código de Productor
// 4: Asegurado 5: Fecha Denuncia (sin uso) 6: Fecha Ocurrencia
// 7: Estado 8: Tipo de Siniestro 9: N° de Siniestro 10: Causa
//
// Por cada fila:
// - Si el N° de Siniestro NO tiene tarea creada todavía: se crea una
//   tarea Tipo="Siniestro", con Fecha de Creación = Fecha de Ocurrencia
//   (no la fecha en que se corre la importación). Compañía siempre
//   "Federación Patronal". Responsable: "Matias" si el código de
//   Productor es 21438, si no "Pilar". El Ramo sale del código antes de
//   la primera "/" en la Póliza, cruzado contra la columna FP de AUX. El
//   cliente se vincula por nombre (Asegurado) contra CLIENTES.
// - Si el N° de Siniestro YA tiene una tarea creada (pasa seguido, porque
//   los siniestros se importan una vez por semana o dos, no en el momento
//   en que se cargan): no se crea una tarea nueva ni se toca su Estado,
//   Responsable o Vencimiento — se le agrega un comentario nuevo arriba
//   de los que ya tenía, igual que "Agregar Comentario" en la pantalla,
//   para no perder el historial de seguimiento que se haya cargado a mano
//   mientras tanto.
//
// Todas las actualizaciones a tareas ya existentes se acumulan en memoria
// y se escriben en una sola operación al final — con reportes grandes,
// la mayoría de las filas suelen ser siniestros que ya estaban cargados,
// así que evitar una escritura por fila importa para no volver a toparnos
// con timeouts.

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

/**
 * Igual formato que usa "Agregar Comentario" en la pantalla ("13/8/26
 * 15:15hs"): día y mes sin cero adelante, año en 2 dígitos, hora sin cero
 * adelante, minutos con cero adelante. Definida acá, reutilizada también
 * por SiniestrosPS.gs.
 */
function formatearFechaComentarioBackend(fecha) {
  const dia = fecha.getDate();
  const mes = fecha.getMonth() + 1;
  const anio = fecha.getFullYear().toString().slice(-2);
  const horas = fecha.getHours();
  const minutos = ("0" + fecha.getMinutes()).slice(-2);
  return dia + "/" + mes + "/" + anio + " " + horas + ":" + minutos + "hs";
}

/**
 * Traduce el código de Productor del CSV a un nombre legible. Si aparece
 * un código que todavía no conocemos, mostramos el código tal cual en vez
 * de perderlo — así se nota que hay que agregarlo acá.
 */
function nombreProductorSiniestrosFP(codigo) {
  if (codigo === "21438") return "Matias";
  if (codigo === "35520") return "Mariana";
  return codigo;
}

function registrarImportacionSiniestrosFP(ss, nombresArchivos, cantidadNuevas, cantidadActualizadas, minFecha, maxFecha, productor) {
  let hoja = ss.getSheetByName("REGISTRO_SINIESTROS_FP");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_SINIESTROS_FP");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Archivo(s)", "Tareas Nuevas", "Tareas Actualizadas", "Ocurrencia Más Antigua", "Ocurrencia Más Reciente", "Productor"]);
  }
  hoja.appendRow([new Date(), nombresArchivos.join(", "), cantidadNuevas, cantidadActualizadas, minFecha, maxFecha, productor || ""]);
  const ultimaFila = hoja.getLastRow();
  hoja.getRange(ultimaFila, 1, 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
  hoja.getRange(ultimaFila, 5, 1, 2).setNumberFormat("dd/mm/yyyy");
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
    const desde = (fila[4] instanceof Date) ? Utilities.formatDate(fila[4], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    const hasta = (fila[5] instanceof Date) ? Utilities.formatDate(fila[5], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";
    return {
      fecha: fecha,
      archivos: fila[1] || "",
      cantidad: fila[2] || 0,
      actualizadas: fila[3] || 0,
      desde: desde,
      hasta: hasta,
      productor: fila[6] || ""
    };
  }).reverse();
}

function procesarSiniestrosFP() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) throw new Error("No se encontró la hoja TAREAS");

  const mapaRamoFP = obtenerMapaRamoFP(); // ya definida en VencimientosFP.gs
  const mapaClientes = obtenerMapaClientesPorNombre(ss); // ya definida en VencimientosRIV.gs

  const datosTareas = hojaTareas.getDataRange().getValues();
  const totalFilasOriginales = datosTareas.length;

  // N° de Siniestro que ya tienen tarea creada -> índice dentro de datosTareas
  const siniestrosYaCreados = new Set();
  const filaPorSiniestro = {};
  for (let i = 1; i < datosTareas.length; i++) {
    const tipoTarea = (datosTareas[i][3] || "").toString().trim().toLowerCase();
    const compania = (datosTareas[i][2] || "").toString().trim().toLowerCase();
    const numSin = (datosTareas[i][15] || "").toString().trim();
    if (tipoTarea === "siniestro" && compania === "federación patronal" && numSin) {
      siniestrosYaCreados.add(numSin);
      filaPorSiniestro[numSin] = i;
    }
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
  const iterador = folder.getFilesByType(MimeType.CSV);
  const archivosCSV = [];
  while (iterador.hasNext()) archivosCSV.push(iterador.next());

  if (archivosCSV.length === 0) {
    return { exito: true, tareasCreadas: 0, actualizadas: 0, omitidas: 0, sinVincular: 0, mensaje: "No hay archivos nuevos para procesar." };
  }

  let idTareaActual = 0;
  if (totalFilasOriginales > 1) {
    idTareaActual = Number(datosTareas[totalFilasOriginales - 1][1]) || 0;
  }

  const filasNuevas = [];
  let filasOmitidas = 0;
  let filasActualizadas = 0;
  let sinVincular = 0;
  let minFecha = null;
  let maxFecha = null;
  const nombresArchivos = [];
  const codigosProductorVistos = new Set();
  let huboActualizacionesEnMemoria = false;

  archivosCSV.forEach(file => {
    nombresArchivos.push(file.getName());
    const csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));

    csvData.forEach(fila => {
      if (!fila || fila.length < 11) return;

      const poliza = (fila[2] || "").toString().trim();
      const productor = (fila[3] || "").toString().trim();
      if (productor) codigosProductorVistos.add(productor);
      const asegurado = (fila[4] || "").toString().trim();
      const fechaOcurrenciaRaw = (fila[6] || "").toString().trim(); // formato yyyy-mm-dd
      const tipoSiniestro = (fila[8] || "").toString().trim();
      const numSiniestro = (fila[9] || "").toString().trim();
      const causa = (fila[10] || "").toString().trim();

      if (!numSiniestro) {
        filasOmitidas++;
        return;
      }

      const fechaOcurrencia = new Date(fechaOcurrenciaRaw + "T12:00:00");
      const fechaOcurrenciaValida = !isNaN(fechaOcurrencia);

      const codigoRamo = poliza.split("/")[0].trim();
      const ramo = mapaRamoFP[codigoRamo] || "";
      const lineaInfo = "Póliza " + poliza +
        (tipoSiniestro ? " - " + tipoSiniestro : "") +
        (causa ? " - " + causa : "");

      if (siniestrosYaCreados.has(numSiniestro)) {
        // Ya existe: fusionamos un comentario nuevo arriba de los que ya
        // tenía, sin tocar Estado/Responsable/Vencimiento de la tarea.
        const idx = filaPorSiniestro[numSiniestro];
        const lineaConFecha = formatearFechaComentarioBackend(new Date()) + " " + lineaInfo;
        const descripcionActual = (datosTareas[idx][4] || "").toString().trim();
        datosTareas[idx][4] = descripcionActual ? (lineaConFecha + "\n" + descripcionActual) : lineaConFecha;
        huboActualizacionesEnMemoria = true;
        filasActualizadas++;

        if (fechaOcurrenciaValida) {
          if (!minFecha || fechaOcurrencia < minFecha) minFecha = fechaOcurrencia;
          if (!maxFecha || fechaOcurrencia > maxFecha) maxFecha = fechaOcurrencia;
        }
        return;
      }

let responsable;

if (productor === "21438") {
  responsable = "Matias";
} else {
  const ramoUpper = String(ramo || "").toUpperCase().trim();
  
  if (ramoUpper === "AUTOMOVILES" || ramoUpper === "AUTOMOTORES") {
    responsable = "Pilar";
  } else if (ramoUpper === "INTEGRALES") {
    responsable = "Mauro";
  } else {
    responsable = "Mariana";
  }
}
      const nombreNormalizado = normalizarNombreParaMatch(asegurado);
      const idCliente = mapaClientes[nombreNormalizado] || "";
      if (!idCliente) sinVincular++;

      idTareaActual++;
      filasNuevas.push([
        fechaOcurrenciaValida ? fechaOcurrencia : new Date(), // A: Fecha Creación = Fecha de Ocurrencia
        idTareaActual,                  // B: ID Tarea
        "Federación Patronal",           // C: Compañía (siempre)
        "Siniestro",                      // D: Tipo de Tarea
        lineaInfo,                         // E: Descripción
        "",                                // F: Vencimiento — se deja vacío, se carga a mano cuando haga falta
        "Sin leer",                       // G: Estado
        "",                                // H
        "",                                // I
        "Normal",                          // J: Prioridad
        "",                                // K: Adjunto
        idCliente,                         // L: ID Cliente
        "Sistema",                         // M: Usuario
        responsable,                       // N: Responsable
        ramo,                              // O: Ramo
        numSiniestro                       // P: N° Siniestro
      ]);

      siniestrosYaCreados.add(numSiniestro); // por si el mismo siniestro aparece 2 veces en el mismo archivo

      if (fechaOcurrenciaValida) {
        if (!minFecha || fechaOcurrencia < minFecha) minFecha = fechaOcurrencia;
        if (!maxFecha || fechaOcurrencia > maxFecha) maxFecha = fechaOcurrencia;
      }
    });

    moverArchivoSiniestrosFPAProcesado(file, folder);
  });

  // Primero volcamos las actualizaciones a tareas existentes (todas juntas)
  if (huboActualizacionesEnMemoria) {
    hojaTareas.getRange(1, 1, totalFilasOriginales, datosTareas[0].length).setValues(datosTareas);
  }

  // Recién después agregamos las tareas realmente nuevas al final
  if (filasNuevas.length > 0) {
    const filaInicio = hojaTareas.getLastRow() + 1;
    hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 16).setValues(filasNuevas);
  }

  if (filasNuevas.length > 0 || filasActualizadas > 0) {
    const productorLabel = Array.from(codigosProductorVistos)
      .map(codigo => nombreProductorSiniestrosFP(codigo))
      .join(", ");
    registrarImportacionSiniestrosFP(ss, nombresArchivos, filasNuevas.length, filasActualizadas, minFecha, maxFecha, productorLabel);
  }

  return {
    exito: true,
    tareasCreadas: filasNuevas.length,
    actualizadas: filasActualizadas,
    omitidas: filasOmitidas,
    sinVincular: sinVincular
  };
}