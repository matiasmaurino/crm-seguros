/**
 * A diferencia de sheet.getLastRow() (que mira TODA la hoja, incluida la
 * columna M "AUXILIAR_BUSQUEDA" que puede tener fórmulas arrastradas mucho
 * más abajo que los datos reales), esta función busca puntualmente la
 * última fila con contenido real en la columna A. Evita que la próxima
 * tanda de datos arranque en cualquier fila lejana dejando un hueco vacío.
 */
function obtenerUltimaFilaConDatosFP(hoja) {
  const ultimaFilaTeorica = hoja.getLastRow();
  if (ultimaFilaTeorica === 0) return 0;

  const valores = hoja.getRange(1, 1, ultimaFilaTeorica, 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    const valor = valores[i][0];
    if (valor !== "" && valor !== null && valor !== undefined) {
      return i + 1;
    }
  }
  return 0;
}

/**
 * Reintenta una operación hasta 3 veces con espera creciente entre intentos,
 * para absorber errores transitorios del servicio de Sheets ("Service
 * Spreadsheets failed...") que no tienen que ver con permisos, solo con
 * que el servicio de Google tuvo un hipo momentáneo.
 */
function ejecutarConReintentos(funcion, intentos) {
  intentos = intentos || 3;
  for (let i = 0; i < intentos; i++) {
    try {
      return funcion();
    } catch (e) {
      if (i === intentos - 1) throw e;
      Logger.log("Reintentando tras error transitorio: " + e.toString());
      Utilities.sleep(2000 * (i + 1)); // 2s, 4s, 6s...
    }
  }
}

/**
 * Lee la columna L ("ID_PROCESAMIENTO", ahora usada para el nombre del
 * archivo de origen) de la hoja "FP" y devuelve el Set de nombres de
 * archivo que ya están cargados, para no reprocesar ninguno.
 */
function obtenerNombresYaImportadosFP(sheet) {
  const ultimaFila = obtenerUltimaFilaConDatosFP(sheet);
  if (ultimaFila < 2) return new Set();
  const valores = sheet.getRange(2, 12, ultimaFila - 1, 1).getValues(); // Columna L
  return new Set(valores.map(fila => fila[0]).filter(v => v));
}

/**
 * Mueve un archivo ya procesado a una subcarpeta "Procesados" dentro de la
 * misma carpeta de origen, para que la próxima corrida no lo vuelva a leer.
 */
function moverArchivoAProcesados(file, folderOrigen) {
  ejecutarConReintentos(() => {
    const nombreSubcarpeta = "Procesados";
    const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
    const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);

    carpetaProcesados.addFile(file);
    folderOrigen.removeFile(file);
  });
}

/**
 * Consolida los CSV de Federación Patronal en la hoja "FP".
 *
 * Cada CSV mensual trae solo movimientos nuevos, así que la función AGREGA
 * las filas del archivo nuevo al final de lo que ya había en "FP", sin
 * tocar lo existente.
 *
 * Control de duplicados: la columna L ("ID_PROCESAMIENTO") ahora guarda el
 * NOMBRE del archivo de origen de cada fila (ya no un timestamp). Antes de
 * procesar, se lee esa columna completa para armar la lista de archivos ya
 * importados, y se salta cualquier archivo cuyo nombre ya esté ahí. Además,
 * cada archivo procesado se mueve a la subcarpeta "Procesados" para que ni
 * siquiera aparezca en la próxima búsqueda.
 */
function consolidarYLimpiarFP() {
  const folderId = '1MFWeyrluXJdDA8pJyuzAGAeRRAOeIHbV';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. CARGAR DICCIONARIO DE CLIENTES
  const hojaClientes = ss.getSheetByName("CLIENTES FP");
  const clientesMap = {};
  if (hojaClientes) {
    const dataClientes = hojaClientes.getDataRange().getValues();
    for (let i = 1; i < dataClientes.length; i++) {
      let nombreCliente = dataClientes[i][2] ? dataClientes[i][2].toString().trim().toUpperCase() : "";
      let idFedPatronal = dataClientes[i][0];
      clientesMap[nombreCliente] = idFedPatronal;
    }
  }

  let nombreHoja = "FP";
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) sheet = ss.insertSheet(nombreHoja);

  const folder = DriveApp.getFolderById(folderId);
  const columnasInteres = [2, 3, 6, 7, 8, 16, 17, 18];

  const nombresImportados = obtenerNombresYaImportadosFP(sheet);

  // Juntamos primero todos los archivos a procesar, salteando los que ya
  // figuran en la columna L de "FP" (por si quedaron sin moverse).
  const archivosCSV = [];
  const iterador = folder.getFilesByType(MimeType.CSV);
  while (iterador.hasNext()) {
    const f = iterador.next();
    if (nombresImportados.has(f.getName())) {
      // Ya está cargado pero seguía en la carpeta: lo movemos igual a
      // Procesados, sin volver a leer sus datos.
      moverArchivoAProcesados(f, folder);
      continue;
    }
    archivosCSV.push(f);
  }

  if (archivosCSV.length === 0) {
    try {
      SpreadsheetApp.getUi().alert("ℹ️ No hay archivos CSV nuevos para procesar en la carpeta de Federación Patronal.");
    } catch (e) {
      console.log("No hay archivos CSV nuevos para procesar (FP).");
    }
    return;
  }

  // Seguimos escribiendo DESPUÉS de lo que ya había, no desde la fila 1
  // (y no desde donde diga getLastRow(), que puede estar inflado por la
  // columna M) — desde la última fila con datos reales en la columna A.
  const ultimaFilaConDatos = obtenerUltimaFilaConDatosFP(sheet);
  let currentRow = ultimaFilaConDatos + 1;
  const hojaEstabaVacia = (ultimaFilaConDatos === 0);

  archivosCSV.forEach((file, index) => {
    const incluirEncabezado = hojaEstabaVacia && index === 0;
    const nombreArchivo = file.getName();

    let csvData = Utilities.parseCsv(file.getBlob().getDataAsString('ISO-8859-1'));
    let datosProcesados = csvData.map((fila, filaIndex) => {
      if (filaIndex === 0 && !incluirEncabezado) return null;

      let filaNueva = columnasInteres.map((idx, colPos) => {
        let valor = fila[idx] ? fila[idx].toString().trim() : "";

        valor = valor.replace(/Ã³/g, "ó").replace(/Ã/g, "Ñ").replace(/[^\x20-\x7E\xC0-\xFF/]/g, "");

        if (colPos === 2 && valor.includes("undefined")) {
          valor = valor.split('/').pop();
        }

        if (colPos === 1 && valor.includes("-")) {
          let partes = valor.substring(0, 10).split('-');
          if (partes.length === 3) return partes[2] + "/" + partes[1] + "/" + partes[0];
        }

        if (colPos >= 4) {
          if (filaIndex === 0) return valor;
          let num = parseFloat(valor.replace(',', ''));
          return isNaN(num) ? valor : Math.trunc(num);
        }
        return valor;
      });

      if (filaIndex === 0) {
        filaNueva.push("ID FedPatronal");
        filaNueva.push("CIA");
        filaNueva.push("PAS AGRUPADO");
        filaNueva.push("ID_PROCESAMIENTO"); // el header no cambia, aunque ahora guarda el nombre del archivo
      } else {
        let nombreAsegurado = filaNueva[2] ? filaNueva[2].toString().trim().toUpperCase() : "";
        let encontrado = clientesMap[nombreAsegurado] || "";
        if (encontrado && typeof encontrado === "string") {
          encontrado = encontrado.replace("M-", "");
        }
        filaNueva.push(encontrado);

        let tieneDato = filaNueva[0] !== "" && filaNueva[0] != null;
        filaNueva.push(tieneDato ? "FP" : "");

        let nombreProductor = filaNueva[0] ? filaNueva[0].toString().trim().toUpperCase() : "";
        if (nombreProductor === "MAURIÑO MATIAS") {
          filaNueva.push("MATIAS");
        } else {
          filaNueva.push(tieneDato ? "DGM" : "");
        }

        // Columna L: ahora el nombre del archivo de origen, no un timestamp
        filaNueva.push(tieneDato ? nombreArchivo : "");
      }
      return filaNueva;
    }).filter(fila => fila !== null);

    if (datosProcesados.length > 0) {
      ejecutarConReintentos(() => {
        sheet.getRange(currentRow, 1, datosProcesados.length, datosProcesados[0].length).setValues(datosProcesados);
      });
      currentRow += datosProcesados.length;
    }

    moverArchivoAProcesados(file, folder);
  });

  // FORMATOS VISUALES
  const ultimaFila = obtenerUltimaFilaConDatosFP(sheet);
  if (ultimaFila > 1) {
    sheet.getRange(2, 2, ultimaFila - 1).setNumberFormat("dd/mm/yyyy");
    sheet.getRange(2, 5, ultimaFila - 1, 4).setNumberFormat("#,##0");
    sheet.autoResizeColumns(1, 12);
  }

  // Sincronización automática al CRM
  sincronizarNuevosAlCRM();

  try {
    SpreadsheetApp.getUi().alert("✅ Procesado correctamente Federación Patronal (" + archivosCSV.length + " archivo/s nuevo/s)");
  } catch (e) {
    console.log("Proceso terminado: FP consolidado (" + archivosCSV.length + " archivo/s nuevo/s).");
  }
}