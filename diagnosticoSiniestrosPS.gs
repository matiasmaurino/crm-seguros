/**
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