/**
 * EJECUTAR MANUALMENTE UNA SOLA VEZ: corrige las filas viejas de
 * REGISTRO_SINIESTROS_FP que quedaron en el formato de 5 columnas (de
 * antes de agregar "Tareas Actualizadas"), insertándoles un 0 en esa
 * columna y corriendo el resto un lugar, para que coincidan con las filas
 * nuevas de 6 columnas.
 *
 * Detecta el formato viejo mirando si la columna D (índice 3) es una
 * fecha — en el formato nuevo esa columna es un número (Tareas
 * Actualizadas), nunca una fecha.
 */
function migrarRegistroSiniestrosFPA6Columnas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("REGISTRO_SINIESTROS_FP");
  if (!hoja) {
    Logger.log("No se encontró la hoja REGISTRO_SINIESTROS_FP.");
    return;
  }

  const data = hoja.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("No hay filas de datos para migrar.");
    return;
  }

  // Reescribimos el encabezado correcto (6 columnas)
  hoja.getRange(1, 1, 1, 6).setValues([[
    "Fecha de Importación", "Archivo(s)", "Tareas Nuevas", "Tareas Actualizadas",
    "Ocurrencia Más Antigua", "Ocurrencia Más Reciente"
  ]]);

  let corregidas = 0;
  const filasCorregidas = data.slice(1).map(fila => {
    const esFormatoViejo = fila[3] instanceof Date;
    if (esFormatoViejo) {
      corregidas++;
      // Viejo: [Fecha, Archivo(s), Cantidad, minFecha, maxFecha]
      // Nuevo: [Fecha, Archivo(s), Nuevas, Actualizadas(0), minFecha, maxFecha]
      return [fila[0], fila[1], fila[2], 0, fila[3], fila[4]];
    }
    return [fila[0], fila[1], fila[2], fila[3], fila[4], fila[5]];
  });

  hoja.getRange(2, 1, filasCorregidas.length, 6).setValues(filasCorregidas);
  hoja.getRange(2, 1, filasCorregidas.length, 1).setNumberFormat("dd/mm/yyyy hh:mm");
  hoja.getRange(2, 5, filasCorregidas.length, 2).setNumberFormat("dd/mm/yyyy");

  Logger.log("Filas corregidas (formato viejo -> nuevo): " + corregidas + " de " + filasCorregidas.length + " totales.");
}/**
 * FUNCIÓN TEMPORAL DE DIAGNÓSTICO — se puede borrar después de usarla.
 * Busca el .xls de Siniestros PS (en la carpeta principal o ya movido a
 * "Procesados"), lo lee con leerFilasDeXLSX() y muestra en el log,
 * fila por fila, exactamente qué valores está encontrando — así vemos
 * dónde se está cortando el procesamiento.
 */



/**
 * FUNCIÓN TEMPORAL DE DIAGNÓSTICO — se puede borrar después de usarla.
 *
 * Busca el último mail de narrowd@pseguros.com.ar con asunto "Detalle de
 * Siniestros Denunciados", saca el adjunto .xls directo del mail (sin
 * pasar por subirlo a Drive primero) e intenta convertirlo. Sirve para
 * confirmar si el problema de conversión es algo específico del camino
 * "subir por la web", o si es el archivo en sí el que Google no puede
 * convertir bien sin importar de dónde venga.
 */

function diagnosticoSiniestrosPS() {
  const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);

  let archivo = null;
  let origen = "";

  // Primero busca en la carpeta principal
  const iterPrincipal = folder.getFiles();
  while (iterPrincipal.hasNext()) {
    const f = iterPrincipal.next();
    if (f.getName().toLowerCase().endsWith('.xls') || f.getName().toLowerCase().endsWith('.xlsx')) {
      archivo = f;
      origen = "carpeta principal";
      break;
    }
  }

  // Si no está ahí, busca en "Procesados"
  if (!archivo) {
    const subcarpetas = folder.getFoldersByName("Procesados");
    if (subcarpetas.hasNext()) {
      const carpetaProcesados = subcarpetas.next();
      const iterProcesados = carpetaProcesados.getFiles();
      let masReciente = null;
      while (iterProcesados.hasNext()) {
        const f = iterProcesados.next();
        if (f.getName().toLowerCase().endsWith('.xls') || f.getName().toLowerCase().endsWith('.xlsx')) {
          if (!masReciente || f.getLastUpdated() > masReciente.getLastUpdated()) masReciente = f;
        }
      }
      if (masReciente) {
        archivo = masReciente;
        origen = "carpeta Procesados";
      }
    }
  }

  if (!archivo) {
    Logger.log("No se encontró ningún .xls/.xlsx en la carpeta principal ni en Procesados.");
    return;
  }

  Logger.log("Archivo encontrado: " + archivo.getName() + " (en " + origen + ")");
  Logger.log("Tipo MIME guardado en Drive: " + archivo.getMimeType());
  Logger.log("Tamaño del archivo: " + archivo.getSize() + " bytes");

  const datos = leerFilasDeXLSX(archivo);
  Logger.log("Cantidad de filas leídas: " + datos.length);

  datos.forEach((fila, i) => {
    Logger.log("Fila " + i + " (largo " + fila.length + "): " + JSON.stringify(fila));
    Logger.log("  -> fila[1] (Siniestro) es: " + JSON.stringify(fila[1]) + " | tipo: " + typeof fila[1]);
  });
}

/**
 * FUNCIÓN TEMPORAL DE DIAGNÓSTICO — se puede borrar después de usarla.
 *
 * Busca el último mail de narrowd@pseguros.com.ar con asunto "Detalle de
 * Siniestros Denunciados", saca el adjunto .xls directo del mail (sin
 * pasar por subirlo a Drive primero) e intenta convertirlo. Sirve para
 * confirmar si el problema de conversión es algo específico del camino
 * "subir por la web", o si es el archivo en sí el que Google no puede
 * convertir bien sin importar de dónde venga.
 */
/**
 * FUNCIÓN TEMPORAL DE DIAGNÓSTICO — se puede borrar después de usarla.
 *
 * Busca el último mail de narrowd@pseguros.com.ar con asunto "Detalle de
 * Siniestros Denunciados", saca el adjunto .xls directo del mail (sin
 * pasar por subirlo a Drive primero) e intenta convertirlo. Sirve para
 * confirmar si el problema de conversión es algo específico del camino
 * "subir por la web", o si es el archivo en sí el que Google no puede
 * convertir bien sin importar de dónde venga.
 */
/**
 * FUNCIÓN TEMPORAL DE DIAGNÓSTICO — se puede borrar después de usarla.
 *
 * Busca el último mail de narrowd@pseguros.com.ar con asunto "Detalle de
 * Siniestros Denunciados", saca el adjunto .xls directo del mail (sin
 * pasar por subirlo a Drive primero) e intenta convertirlo. Sirve para
 * confirmar si el problema de conversión es algo específico del camino
 * "subir por la web", o si es el archivo en sí el que Google no puede
 * convertir bien sin importar de dónde venga.
 */
function diagnosticoSiniestrosPSDesdeEmail() {
  const query = 'from:narrowd@pseguros.com.ar subject:"Detalle de Siniestros Denunciados"';
  const threads = GmailApp.search(query, 0, 5);

  if (threads.length === 0) {
    Logger.log("No se encontró ningún mail que matchee esa búsqueda. Revisá el remitente/asunto.");
    return;
  }

  const mensajes = threads[0].getMessages();
  const ultimoMensaje = mensajes[mensajes.length - 1];
  Logger.log("Mail encontrado: \"" + ultimoMensaje.getSubject() + "\" del " + ultimoMensaje.getDate());

  const adjuntos = ultimoMensaje.getAttachments();
  Logger.log("Cantidad de adjuntos en ese mail: " + adjuntos.length);

  let adjuntoXLS = null;
  adjuntos.forEach(a => {
    Logger.log(" - " + a.getName() + " | tipo: " + a.getContentType() + " | " + a.getSize() + " bytes");
    if (a.getName().toLowerCase().endsWith('.xls') || a.getName().toLowerCase().endsWith('.xlsx')) {
      adjuntoXLS = a;
    }
  });

  if (!adjuntoXLS) {
    Logger.log("No se encontró ningún adjunto .xls/.xlsx en ese mail.");
    return;
  }

  Logger.log("--- Probando convertir el adjunto directo del mail (sin pasar por Drive primero) ---");

  try {
    // Gmail reporta este adjunto con un tipo MIME no estándar
    // (application/x-msexcel); lo forzamos al tipo correcto antes de
    // pedirle a Drive que lo convierta.
    const blob = adjuntoXLS.copyBlob().setContentType('application/vnd.ms-excel');
    Logger.log("Tipo MIME del blob justo antes de mandarlo a Drive: " + blob.getContentType());

    // Usamos el string literal en vez de MimeType.GOOGLE_SHEETS, para
    // descartar cualquier problema con esa constante puntual.
    const resource = { title: "temp_test_email", mimeType: "application/vnd.google-apps.spreadsheet" };
    const tempFile = Drive.Files.insert(resource, blob, { convert: true });

    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    const datos = tempSheet.getDataRange().getValues();

    Logger.log("Filas leídas: " + datos.length);
    datos.forEach((fila, i) => {
      Logger.log("Fila " + i + " (largo " + fila.length + "): " + JSON.stringify(fila));
    });

    Drive.Files.remove(tempFile.id);
  } catch (e) {
    Logger.log("Error al convertir: " + e.toString());
  }
}