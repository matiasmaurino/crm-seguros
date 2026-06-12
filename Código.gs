
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <-- ACÁ VA LA LÍNEA
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- GESTIÓN DE CLIENTES ---

function obtenerListaClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(r => {
    let etiquetas = [];
    
    // Revisamos cada columna por separado para armar las etiquetas visuales
    if (r[6]) etiquetas.push("RIV");      // Columna G: Rivadavia
    if (r[7]) etiquetas.push("PS");       // Columna H: Provincia Seguros
    if (r[8]) etiquetas.push("FP");       // Columna I: Federación Patronal

    let etiquetaFinal = etiquetas.length > 0 ? " [" + etiquetas.join(" / ") + "]" : "";

    return {
      id: r[0],                            // Columna A
      nombre: r[1] + etiquetaFinal,        // Nombre con corchetes para buscador
      nombrePuro: r[1],                    // Columna B (Nombre original)
      dni: r[2],                           // <--- CORREGIDO: Cambiado 'cuit' por 'dni' para alinearse con Scripts.html
      domicilio: r[3],                     // Columna D
      telefono: r[4],                      // Columna E
      email: r[5],                         // Columna F
      rivadavia: r[6],                     // Columna G
      provincia: r[7],                     // Columna H
      fedPatronal: r[8],                   // Columna I
      relacionados: r[9],                  // Columna J
      observaciones: r[10]                 // Columna K
    };
  });
}
function guardarCliente(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  const data = hoja.getDataRange().getValues();
  let idClienteActual = datos.id;
  let filaIndex = -1;

  if (idClienteActual) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == idClienteActual) {
        filaIndex = i + 1;
        break;
      }
    }
  }

  if (filaIndex === -1) {
    filaIndex = hoja.getLastRow() + 1;
    idClienteActual = data.length > 1 ? Number(data[data.length-1][0]) + 1 : 1;
    hoja.getRange(filaIndex, 1).setValue(idClienteActual);
  }

  const relacionadosJSON = JSON.stringify(datos.relacionados || []);

  hoja.getRange(filaIndex, 2, 1, 9).setValues([[
    datos.nombre.toUpperCase(), 
    datos.dni, 
    datos.domicilio, 
    datos.telefono, 
    datos.email, 
    datos.rivadavia,
    datos.provincia,
    datos.fedPatronal,
    relacionadosJSON
  ]]);

  // Sincronización de relacionados (vínculo doble)
  if (datos.relacionados && datos.relacionados.length > 0) {
    const todos = hoja.getDataRange().getValues();
    datos.relacionados.forEach(rel => {
      for (let j = 1; j < todos.length; j++) {
        if (todos[j][0] == rel.id) {
          let susRel = [];
          try { susRel = JSON.parse(todos[j][9] || "[]"); } catch(e) { susRel = []; }
          if (!susRel.some(x => x.id == idClienteActual)) {
            susRel.push({ id: idClienteActual, nombre: datos.nombre, rol: rel.rol });
            hoja.getRange(j + 1, 10).setValue(JSON.stringify(susRel));
          }
          break;
        }
      }
    });
  }
  return { exito: true, mensaje: "Cliente guardado con éxito" };
}

// --- GESTIÓN DE TAREAS ---

function obtenerTareas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const clientes = obtenerListaClientes();
  
  return data.slice(1).map((r, index) => {
    const idClienteEnFila = r[11]; 
    const c = clientes.find(x => x.id == idClienteEnFila);
    
    let fCreacionRaw = r[0];
    let fCreacionFormat = (fCreacionRaw instanceof Date) ? Utilities.formatDate(fCreacionRaw, "GMT-3", "dd/MM/yy") : "-";

    let fechaRaw = r[5]; 
    let fechaFormateada = "";
    let fechaISO = "";

    if (fechaRaw instanceof Date && !isNaN(fechaRaw)) {
      fechaFormateada = Utilities.formatDate(fechaRaw, "GMT-3", "dd/MM/yy");
      fechaISO = Utilities.formatDate(fechaRaw, "GMT-3", "yyyy-MM-dd");
    }

    return {
      id_fila: index + 2,
      idCliente: idClienteEnFila,
      clienteNombre: c ? c.nombre : "Sin asignar",
      creacion: fCreacionFormat,
      compania: r[2],    
      tipoTarea: r[3],   
      descripcion: r[4], 
      vencimiento: fechaISO,
      vencimientoFormat: fechaFormateada,
      estado: r[6],      
      prioridad: r[9],   
      adjunto: r[10],
      responsable: r[13],
      ramo: r[14] || "",    // <-- NUEVO: Lee la columna O (índice 14)
      usuario: r[12] || "-" // Columna M
    };
  }).filter(t => t.compania !== "").reverse(); 
}

function guardarTarea(t, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  const data = hoja.getDataRange().getValues();
  let idTareaActual = null;

  if (t.id_fila) {
    // Si estamos EDITANDO una fila existente, no tocamos el ID, mantenemos el que ya tiene la fila
    // El ID real de la tarea está en la columna B (índice 1 de la fila)
    idTareaActual = data[Number(t.id_fila) - 1][1];
    
    // CORRECCIÓN: Volvemos a expandir el rango a 13 columnas (Columna C hasta la O)
    // Agregamos los dos campos vacíos "" originales correspondientes a las columnas H e I para que mantenga la estructura de la base de datos
    hoja.getRange(Number(t.id_fila), 3, 1, 13).setValues([[
      t.compania,                    // C
      t.tipoTarea,                   // D
      t.descripcion,                 // E
      t.vencimiento ? new Date(t.vencimiento + "T12:00:00") : "", // F
      t.estado,                      // G
      "",                            // H (Vacío estructural)
      "",                            // I (Vacío estructural)
      t.prioridad,                   // J
      t.adjunto,                     // K
      t.idCliente,                   // L
      usuarioActivo || "Sistema",    // M: Usuario Logueado
      t.responsable || "",           // N: Responsable seleccionado
      t.ramo || ""                   // O: Ramo seleccionado
    ]]);
  } else {
    // Si es una TAREA NUEVA, calculamos el ID de forma correlativa
    idTareaActual = data.length > 1 ? Number(data[data.length - 1][1]) + 1 : 1;
    
    if (isNaN(idTareaActual)) {
      idTareaActual = 1;
    }

    // Preparamos los valores para la fila nueva (Columna A a O)
    const filaValores = [
      new Date(),     // A: Fecha Creación
      idTareaActual,  // B: ID Tarea
      t.compania,     // C
      t.tipoTarea,    // D
      t.descripcion,  // E
      t.vencimiento ? new Date(t.vencimiento + "T12:00:00") : "", // F
      t.estado,      // G
      "",            // H
      "",            // I
      t.prioridad,   // J
      t.adjunto,     // K
      t.idCliente,   // L
      usuarioActivo || "Sistema", // M: Usuario Logueado
      t.responsable || "",         // N: Responsable seleccionado
      t.ramo || ""                // O: Ramo seleccionado
    ];

    hoja.appendRow(filaValores);
    
    // --- LÓGICA DEL CALENDARIO ---
    if (t.vencimiento) {
      try {
        agendarTareaEnCalendar(t.idCliente, t.tipoTarea, t.vencimiento);
      } catch (e) {
        console.log("Error al crear evento en calendario: " + e.toString());
      }
    }
  }
}
// --- ARCHIVOS Y LOGIN ---

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

function validarLogin(usuario, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("LOGIN");
    if (!hoja) return { exito: false, mensaje: "Hoja LOGIN no encontrada" };
    
    const datos = hoja.getDataRange().getValues();
    
    // Comparamos usuario y contraseña (limpiando espacios)
    for (let i = 1; i < datos.length; i++) {
      let userDb = datos[i][0].toString().trim();
      let passDb = datos[i][1].toString().trim();
      
      if (userDb === usuario.toString().trim() && passDb === password.toString().trim()) {
        return { exito: true, usuario: userDb };
      }
    }
    return { exito: false, mensaje: "Usuario o contraseña incorrectos" };
  } catch(e) {
    return { exito: false, mensaje: "Error: " + e.toString() };
  }
}

/**
 * CREA EL EVENTO EN TU CALENDARIO ESPECÍFICO (CORREGIDO)
 */
function agendarTareaEnCalendar(idCliente, nombreTarea, fechaVencimiento) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaClientes = ss.getSheetByName("CLIENTES");
    const datosClientes = hojaClientes.getDataRange().getValues();
    let nombre = "Cliente desconocido";
    let telefono = "No cargado";
    let email = "";

    // 1. Buscar los datos del cliente por su ID
    for (let i = 1; i < datosClientes.length; i++) {
      if (datosClientes[i][0].toString() === idCliente.toString()) { 
        nombre = datosClientes[i][1];   // Columna B: Nombre
        telefono = datosClientes[i][4]; // Columna E: Teléfono
        email = datosClientes[i][5];    // Columna F: Email
        break;
      }
    }

    // --- FORZAMOS TU CALENDARIO PRINCIPAL ---
    const idCalendarioFijo = "matiasmaurino@gmail.com";
    const calendario = CalendarApp.getCalendarById(idCalendarioFijo);
    
    if (!calendario) {
      console.log("No se pudo abrir el calendario: " + idCalendarioFijo + ". Revisar permisos.");
      return;
    }

    const titulo = "Tarea: " + nombreTarea + " - " + nombre;
    
    // 2. Configurar horario (9:00 AM)
    const fechaEvento = new Date(fechaVencimiento + "T09:00:00");
    const finEvento = new Date(fechaEvento.getTime() + 60 * 60 * 1000); // 1 hora de duración

    const descripcion = "ID Cliente: " + idCliente +
                        "\nNombre: " + nombre + 
                        "\nTarea: " + nombreTarea + 
                        "\nTeléfono: " + telefono + 
                        "\nEmail: " + email;

    const evento = calendario.createEvent(titulo, fechaEvento, finEvento, {
      description: descripcion
    });

    // 3. Notificación por correo 1 día antes
    evento.addEmailReminder(1440);
    
    console.log("Evento creado con éxito en " + idCalendarioFijo);
    return evento.getId();
    
  } catch (e) {
    console.log("Error interno en agendarTareaEnCalendar: " + e.toString());
  }
}

function obtenerResponsables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("LOGIN");
  if (!hoja) return [];
  const datos = hoja.getDataRange().getValues();
  // Retorna nombres de columna C (índice 2), sin el encabezado
  return datos.slice(1).map(fila => fila[2]).filter(nombre => nombre);
}
function obtenerRamosLista() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("RAMO");
    if (!hoja) return [];
    
    const data = hoja.getDataRange().getValues();
    if (data.length <= 1) return []; // Si solo está el encabezado o está vacía
    
    // Mapea la columna A (índice 0), omitiendo la fila del encabezado y limpiando vacíos
    return data.slice(1).map(fila => fila[0].toString().trim()).filter(ramo => ramo);
  } catch(e) {
    console.log("Error al obtener ramos: " + e.toString());
    return [];
  }
}