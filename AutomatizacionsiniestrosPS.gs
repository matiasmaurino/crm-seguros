/**
 * Busca el/los mail(es) nuevos de narrowd@pseguros.com.ar con asunto
 * "Detalle de Siniestros Denunciados" que todavía no se procesaron (los
 * marca con la etiqueta "CRM-Procesado" en Gmail una vez que los toma, así
 * la próxima corrida no los vuelve a agarrar). Guarda el adjunto en la
 * carpeta de Drive y dispara procesarSiniestrosPS() automáticamente.
 *
 * Pensada para correr sola con un disparador semanal (ver instrucciones).
 */
function importarSiniestrosPSDesdeEmailAutomatico() {
  const NOMBRE_LABEL = "CRM-Procesado";
  let label = GmailApp.getUserLabelByName(NOMBRE_LABEL);
  if (!label) label = GmailApp.createLabel(NOMBRE_LABEL);

  const query = 'from:narrowd@pseguros.com.ar subject:"Detalle de Siniestros Denunciados" -label:' + NOMBRE_LABEL;
  const threads = GmailApp.search(query, 0, 10);

  if (threads.length === 0) {
    Logger.log("No hay mails nuevos de Siniestros PS para procesar.");
    return;
  }

  const folder = DriveApp.getFolderById(ID_CARPETA_SINIESTROS_FP);
  let archivosSubidos = 0;

  threads.forEach(thread => {
    const mensajes = thread.getMessages();
    mensajes.forEach(mensaje => {
      const adjuntos = mensaje.getAttachments();
      adjuntos.forEach(adjunto => {
        const nombreLower = adjunto.getName().toLowerCase();
        if (nombreLower.endsWith('.xls') || nombreLower.endsWith('.xlsx')) {
          const tipoCorrecto = nombreLower.endsWith('.xlsx')
            ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            : 'application/vnd.ms-excel';
          const blobCorregido = adjunto.copyBlob().setContentType(tipoCorrecto);
          folder.createFile(blobCorregido).setName(adjunto.getName());
          archivosSubidos++;
          Logger.log("Adjunto guardado en Drive: " + adjunto.getName());
        }
      });
    });
    // Lo marcamos como procesado para que la próxima corrida no lo vuelva a tomar
    thread.addLabel(label);
  });

  Logger.log("Total de archivos subidos desde mail: " + archivosSubidos);

  if (archivosSubidos > 0) {
    const resultado = procesarSiniestrosPS();
    Logger.log("Resultado del procesamiento: " + JSON.stringify(resultado));
  }
}