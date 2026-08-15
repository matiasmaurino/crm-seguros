// --- MÓDULO CLIENTES FP ---
// Se sube el CSV del padrón de clientes de Federación Patronal (el mismo
// que hoy se pega a mano en la hoja CLIENTES FP de Liquidaciones). A
// diferencia de Vencimientos FP, acá no alcanza con "agregar lo nuevo":
// un cliente que ya está cargado puede volver a aparecer con datos
// actualizados (teléfono, dirección, etc.), así que se hace un "upsert"
// por Matrícula — si ya existe, se pisa esa fila entera; si es nueva, se
// agrega.
//
// Columnas del CSV (19 en total, A a S):
// Matrícula, Tipo, Apellido y Nombre, Garantiza Cobertura, Sexo, IVA,
// Fecha Nac., Lugar Nacimiento, Estado Civil, Nacionalidad, Dirección, CP,
// CPA, Localidad, Documento, Teléfono, CUIT/CUIL/CDI, E-mail, Cobrador

const ID_CARPETA_CLIENTES_FP = "1IshZBuzRnOSDpayYkVvGgD3-isje95ug";
const COLUMNAS_CLIENTES_FP = 19;

/**
 * Reintenta una operación hasta 3 veces con espera creciente, para
 * absorber errores transitorios del servicio de Sheets. Copia local: la
 * versión original vive en el proyecto de Apps Script de Liquidaciones,
 * que es un proyecto distinto al del CRM y no comparte funciones con él.
 */
function ejecutarConReintentosClientesFP(funcion, intentos) {
  intentos = intentos || 3;
  for (let i = 0; i < intentos; i++) {
    try {
      return funcion();
    } catch (e) {
      if (i === intentos - 1) throw e;
      Logger.log("Reintentando tras error transitorio: " + e.toString());
      Utilities.sleep(2000 * (i + 1));
    }
  }
}

/**
 * Sube el CSV elegido en la web a la carpeta de Drive de Clientes FP.
 */
function subirCSVClientesFP(base64, nombreArchivo) {
  try {
    const folder = DriveApp.getFolderById(ID_CARPETA_CLIENTES_FP);
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, "text/csv", nombreArchivo));
    return { exito: true, idArchivo: archivo.getId(), nombre: archivo.getName() };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

function moverArchivoClientesFPAProcesado(file, folderOrigen) {
  const nombreSubcarpeta = "Procesados";
  const subcarpetas = folderOrigen.getFoldersByName(nombreSubcarpeta);
  const carpetaProcesados = subcarpetas.hasNext() ? subcarpetas.next() : folderOrigen.createFolder(nombreSubcarpeta);
  carpetaProcesados.addFile(file);
  folderOrigen.removeFile(file);
}

/**
 * Registra en REGISTRO_IMPORTACION_CLIENTES_FP (en el mismo archivo de
 * Liquidaciones, se crea sola si no existe) una fila por cada importación,
 * con los archivos procesados y cuántos clientes se agregaron/actualizaron.
 */
function registrarImportacionClientesFP(ss, nombresArchivos, nuevos, actualizados) {
  let hoja = ss.getSheetByName("REGISTRO_IMPORTACION_CLIENTES_FP");
  if (!hoja) {
    hoja = ss.insertSheet("REGISTRO_IMPORTACION_CLIENTES_FP");
  }
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(["Fecha de Importación", "Archivo(s)", "Clientes Nuevos", "Clientes Actualizados"]);
  }
  hoja.appendRow([new Date(), nombresArchivos.join(", "), nuevos, actualizados]);
}

/**
 * Devuelve el historial de importaciones de Clientes FP (más recientes primero).
 */
function obtenerHistorialClientesFP() {
  const ss = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
  const hoja = ss.getSheetByName("REGISTRO_IMPORTACION_CLIENTES_FP");
  if (!hoja) return [];

  const data = hoja.getDataRange().getValues();
  if (data.length === 0) return [];

  const primeraEsEncabezado = !(data[0][0] instanceof Date);
  const filas = primeraEsEncabezado ? data.slice(1) : data;

  return filas.map(fila => {
    const fecha = (fila[0] instanceof Date) ? Utilities.formatDate(fila[0], ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm") : "";
    return {
      fecha: fecha,
      archivos: fila[1] || "",
      nuevos: fila[2] || 0,
      actualizados: fila[3] || 0
    };
  }).reverse();
}

/**
 * Última fila con contenido real en la columna A de una hoja (evita
 * confiar en getLastRow(), que puede quedar inflado por columnas con
 * fórmulas más allá de los datos reales).
 */
function obtenerUltimaFilaRealClientesCRM(hoja) {
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
 * Sincroniza CLIENTES FP con la hoja CLIENTES del CRM, matcheando por CUIT:
 * - Si el cliente ya existe en el CRM, actualiza Domicilio/Teléfono/Email.
 * - Si no existe, lo agrega como cliente nuevo (con esos mismos datos
 *   completos, no en blanco).
 * Se llama automáticamente al final de procesarClientesFP().
 *
 * Todos los cambios (actualizaciones Y altas) se hacen en memoria y se
 * escriben en UNA sola operación al final — antes se escribía cada
 * actualización por separado (una llamada a la hoja por cliente ya
 * existente), que con cientos de clientes ya cargados se volvía muy lento.
 */
function sincronizarClientesFPconCRM(ssCRM, ssLiq) {
  const hojaOrigen = ssLiq.getSheetByName("CLIENTES FP");
  const hojaDestino = ssCRM.getSheetByName("CLIENTES");

  if (!hojaOrigen || !hojaDestino) {
    Logger.log("No se encontró CLIENTES FP o CLIENTES para sincronizar con el CRM.");
    return { nuevosCRM: 0, actualizadosCRM: 0 };
  }

  const datosOrigen = hojaOrigen.getDataRange().getValues();

  const totalFilasDestino = hojaDestino.getLastRow();
  const columnasDestino = Math.max(hojaDestino.getLastColumn(), 9); // al menos hasta la I (FedPatronal)
  const datosDestino = totalFilasDestino > 0
    ? hojaDestino.getRange(1, 1, totalFilasDestino, columnasDestino).getValues()
    : [];

  const encabezadoDestino = datosDestino.length > 0
    ? datosDestino[0]
    : ["ID", "Nombre", "CUIL/CUIT", "Domicilio", "Telefono", "Email", "Rivadavia", "Provincia", "FedPatronal"];
  const filasDestino = datosDestino.length > 1 ? datosDestino.slice(1) : [];

  // Mapa CUIT -> índice dentro de "filasDestino" (en memoria, no fila real de hoja)
  const indicePorCuit = {};
  filasDestino.forEach((fila, idx) => {
    const cuit = fila[2] ? fila[2].toString().replace(/\D/g, "").trim() : "";
    if (cuit.length > 5) indicePorCuit[cuit] = idx;
  });

  let ultimoId = 0;
  filasDestino.forEach(fila => {
    const id = parseInt(fila[0]);
    if (!isNaN(id) && id > ultimoId) ultimoId = id;
  });

  let nuevosCRM = 0;
  let actualizadosCRM = 0;

  for (let i = 1; i < datosOrigen.length; i++) {
    const nombre = datosOrigen[i][2] ? datosOrigen[i][2].toString().trim().toUpperCase() : "";
    const cuit = datosOrigen[i][16] ? datosOrigen[i][16].toString().replace(/\D/g, "").trim() : "";
    const numClienteOrigen = datosOrigen[i][0] ? datosOrigen[i][0].toString().trim() : "";
    const direccion = datosOrigen[i][10] || "";
    const telefono = datosOrigen[i][15] || "";
    const email = datosOrigen[i][17] || "";

    if (nombre === "" || cuit.length < 7 || cuit.length > 11) continue;

    if (indicePorCuit.hasOwnProperty(cuit)) {
      // Ya existe: actualizamos Domicilio/Teléfono/Email en memoria (columnas D, E, F)
      const idx = indicePorCuit[cuit];
      filasDestino[idx][3] = direccion;
      filasDestino[idx][4] = telefono;
      filasDestino[idx][5] = email;
      actualizadosCRM++;
    } else {
      ultimoId++;
      const filaNueva = new Array(encabezadoDestino.length).fill("");
      filaNueva[0] = ultimoId;
      filaNueva[1] = nombre;
      filaNueva[2] = cuit;
      filaNueva[3] = direccion;
      filaNueva[4] = telefono;
      filaNueva[5] = email;
      filaNueva[8] = numClienteOrigen; // columna I: FedPatronal
      filasDestino.push(filaNueva);
      indicePorCuit[cuit] = filasDestino.length - 1;
      nuevosCRM++;
    }
  }

  const resultadoFinal = [encabezadoDestino].concat(filasDestino);
  ejecutarConReintentosClientesFP(() => {
    hojaDestino.getRange(1, 1, resultadoFinal.length, encabezadoDestino.length).setValues(resultadoFinal);
  });

  return { nuevosCRM: nuevosCRM, actualizadosCRM: actualizadosCRM };
}

/**
 * Procesa todos los CSV que estén en la carpeta de Clientes FP: por cada
 * fila, si ya existe un cliente con esa Matrícula en CLIENTES FP, actualiza
 * esa fila entera; si no existe, la agrega. Al terminar, escribe el
 * resultado completo en una sola operación, mueve los archivos ya
 * procesados a "Procesados", y sincroniza los cambios con el CRM (agrega
 * clientes nuevos y actualiza Domicilio/Teléfono/Email de los existentes).
 */
function procesarClientesFP() {
  const ss = SpreadsheetApp.openById(ID_SPREADSHEET_LIQUIDACIONES);
  const hoja = ss.getSheetByName("CLIENTES FP");
  if (!hoja) throw new Error("No se encontró la hoja CLIENTES FP en el archivo de Liquidaciones.");

  const folder = DriveApp.getFolderById(ID_CARPETA_CLIENTES_FP);
  const iterador = folder.getFilesByType(MimeType.CSV);
  const archivosCSV = [];
  while (iterador.hasNext()) archivosCSV.push(iterador.next());

  if (archivosCSV.length === 0) {
    return { exito: true, nuevos: 0, actualizados: 0, mensaje: "No hay archivos nuevos para procesar." };
  }

  const encabezadoPorDefecto = [
    "Matrícula", "Tipo", "Apellido y Nombre", "Garantiza Cobertura", "Sexo", "IVA",
    "Fecha Nac.", "Lugar Nacimiento", "Estado Civil", "Nacionalidad", "Dirección", "CP",
    "CPA", "Localidad", "Documento", "Teléfono", "CUIT/CUIL/CDI", "E-mail", "Cobrador"
  ];

  const totalFilasActuales = hoja.getLastRow();
  const datosActuales = totalFilasActuales > 0
    ? hoja.getRange(1, 1, totalFilasActuales, COLUMNAS_CLIENTES_FP).getValues()
    : [];

  const encabezado = datosActuales.length > 0 ? datosActuales[0] : encabezadoPorDefecto;
  const filas = datosActuales.length > 1 ? datosActuales.slice(1) : [];

  // Mapa Matrícula -> índice dentro de "filas", para poder actualizar en memoria
  const indicePorMatricula = {};
  filas.forEach((fila, idx) => {
    const matricula = (fila[0] || "").toString().trim();
    if (matricula) indicePorMatricula[matricula] = idx;
  });

  let nuevos = 0;
  let actualizados = 0;
  const nombresArchivos = [];

  archivosCSV.forEach(file => {
    nombresArchivos.push(file.getName());
    const csvData = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8'));

    for (let i = 1; i < csvData.length; i++) {
      const filaCsv = csvData[i];
      if (!filaCsv || !filaCsv[0]) continue;

      const matriculaCruda = filaCsv[0].toString().trim();
      if (!matriculaCruda) continue;

      // Sacamos el prefijo "M-" (o "M" sin guion) para dejar solo el número,
      // igual que ya se hace en las liquidaciones de FP.
      const matricula = matriculaCruda.replace(/^M-?/i, "").trim();
      if (!matricula) continue;

      // Nos aseguramos de que la fila tenga exactamente COLUMNAS_CLIENTES_FP columnas
      const filaNormalizada = filaCsv.slice(0, COLUMNAS_CLIENTES_FP);
      while (filaNormalizada.length < COLUMNAS_CLIENTES_FP) filaNormalizada.push("");
      filaNormalizada[0] = matricula; // columna A limpia, sin "M-"

      if (indicePorMatricula.hasOwnProperty(matricula)) {
        filas[indicePorMatricula[matricula]] = filaNormalizada;
        actualizados++;
      } else {
        filas.push(filaNormalizada);
        indicePorMatricula[matricula] = filas.length - 1;
        nuevos++;
      }
    }

    moverArchivoClientesFPAProcesado(file, folder);
  });

  const resultadoFinal = [encabezado].concat(filas);

  // Solo tocamos las columnas A-S: si hubiera alguna columna extra más allá
  // (por ejemplo una fórmula auxiliar), queda intacta.
  if (totalFilasActuales > 0) {
    hoja.getRange(1, 1, totalFilasActuales, COLUMNAS_CLIENTES_FP).clearContent();
  }
  ejecutarConReintentosClientesFP(() => {
    hoja.getRange(1, 1, resultadoFinal.length, COLUMNAS_CLIENTES_FP).setValues(resultadoFinal);
  });

  registrarImportacionClientesFP(ss, nombresArchivos, nuevos, actualizados);

  // Sincronizamos automáticamente con la hoja CLIENTES del CRM: agrega
  // clientes nuevos y actualiza Domicilio/Teléfono/Email de los existentes.
  const ssCRM = SpreadsheetApp.getActiveSpreadsheet();
  let resultadoCRM = { nuevosCRM: 0, actualizadosCRM: 0 };
  try {
    resultadoCRM = sincronizarClientesFPconCRM(ssCRM, ss);
  } catch (e) {
    Logger.log("Error al sincronizar Clientes FP con el CRM: " + e.toString());
  }

  return {
    exito: true,
    nuevos: nuevos,
    actualizados: actualizados,
    nuevosCRM: resultadoCRM.nuevosCRM,
    actualizadosCRM: resultadoCRM.actualizadosCRM
  };
}