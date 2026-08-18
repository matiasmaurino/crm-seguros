// --- PUNTO DE ENTRADA DE LA WEBAPP ---

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- LOGIN ---

function validarLogin(usuario, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("LOGIN");
    if (!hoja) return { exito: false, mensaje: "Hoja LOGIN no encontrada" };
    
    const datos = hoja.getDataRange().getValues();
    
    for (let i = 1; i < datos.length; i++) {
      let userDb = datos[i][0].toString().trim();
      let passDb = datos[i][1].toString().trim();
      
      if (userDb === usuario.toString().trim() && passDb === password.toString().trim()) {
        // Devolvemos el nombre legible (columna C, "Responsable"), no el
        // usuario de login en sí — así queda identificado por su nombre
        // en Usuario/Responsable de TAREAS y en los comentarios automáticos.
        const nombreResponsable = (datos[i][2] || "").toString().trim();
        return { exito: true, usuario: nombreResponsable || userDb };
      }
    }
    return { exito: false, mensaje: "Usuario o contraseña incorrectos" };
  } catch(e) {
    return { exito: false, mensaje: "Error: " + e.toString() };
  }
}

function obtenerResponsables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("LOGIN");
  if (!hoja) return [];
  const datos = hoja.getDataRange().getValues();
  return datos.slice(1).map(fila => fila[2]).filter(nombre => nombre);
}

// --- LISTAS AUXILIARES ---

function obtenerRamosLista() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RAMO");
    if (!hoja) return [];
    
    const data = hoja.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    return data.slice(1).map(fila => fila[0].toString().trim()).filter(ramo => ramo);
  } catch(e) {
    console.log("Error al obtener ramos: " + e.toString());
    return [];
  }
}

// --- ARCHIVOS ---

function subirArchivoADrive(base64, nombre) {
  try {
    const folders = DriveApp.getFoldersByName("ADJUNTOS_CRM");
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("ADJUNTOS_CRM");
    const data = Utilities.base64Decode(base64.split(",")[1]);
    const archivo = folder.createFile(Utilities.newBlob(data, null, nombre));
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch(e) {
    return "Error al subir: " + e.toString();
  }
}