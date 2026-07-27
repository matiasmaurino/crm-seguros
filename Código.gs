// --- MAPEO DE RESPONSABLE -> CALENDARIO ---
const CALENDARIOS_POR_RESPONSABLE = {
  "Matias": "0q14of40e34nhec1lnjh2kami8@group.calendar.google.com"
};
const CALENDARIO_POR_DEFECTO = "matiasmaurino@gmail.com"; // Mariana, Mauro, Pilar y sin asignar caen acá

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

// --- GESTIÓN DE CLIENTES ---

function obtenerListaClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(r => {
    let etiquetas = [];
    
    if (r[6]) etiquetas.push("RIV");
    if (r[7]) etiquetas.push("PS");
    if (r[8]) etiquetas.push("FP");

    let etiquetaFinal = etiquetas.length > 0 ? " [" + etiquetas.join(" / ") + "]" : "";

    return {
      id: r[0],
      nombre: r[1] + etiquetaFinal,
      nombrePuro: r[1],
      dni: r[2],
      domicilio: r[3],
      telefono: r[4],
      email: r[5],
      rivadavia: r[6],
      provincia: r[7],
      fedPatronal: r[8],
      relacionados: r[9],
      observaciones: r[10]
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
      ramo: r[14] || "",
      usuario: r[12] || "-"
    };
  }).filter(t => t.compania !== "").reverse(); 
}

function guardarTarea(t, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  const data = hoja.getDataRange().getValues();
  let idTareaActual = null;

  if (t.id_fila) {
    const filaActual = Number(t.id_fila);
    const filaAnteriorValores = data[filaActual - 1];
    idTareaActual = filaAnteriorValores[1];

    const vencimientoAnterior = filaAnteriorValores[5];
    const idEventoAnterior = (filaAnteriorValores[8] || "").toString().trim(); // columna I
    const nuevoVencimiento = t.vencimiento ? new Date(t.vencimiento + "T12:00:00") : "";

    // Detectamos si cambió la fecha de vencimiento
    let vencimientoCambio = false;
    const vencAnteriorStr = (vencimientoAnterior instanceof Date) ? vencimientoAnterior.toDateString() : "";
    const vencNuevoStr = (nuevoVencimiento instanceof Date) ? nuevoVencimiento.toDateString() : "";
    if (vencAnteriorStr !== vencNuevoStr) vencimientoCambio = true;

    hoja.getRange(filaActual, 3, 1, 13).setValues([[
      t.compania,                    // C
      t.tipoTarea,                   // D
      t.descripcion,                 // E
      nuevoVencimiento,              // F
      t.estado,                      // G
      "",                            // H (vacío estructural)
      vencimientoCambio ? "" : idEventoAnterior, // I: conservamos el ID si la fecha no cambió
      t.prioridad,                   // J
      t.adjunto,                     // K
      t.idCliente,                   // L
      usuarioActivo || "Sistema",    // M
      t.responsable || "",           // N
      t.ramo || ""                   // O
    ]]);

    // Si cambió el vencimiento, borramos el evento viejo y agendamos uno nuevo
    if (vencimientoCambio) {
      if (idEventoAnterior !== "") {
        try {
          const cal = CalendarApp.getCalendarById(CALENDARIO_POR_DEFECTO);
          const eventoViejo = cal ? cal.getEventById(idEventoAnterior) : null;
          if (eventoViejo) eventoViejo.deleteEvent();
        } catch (e) {
          console.log("No se pudo borrar el evento anterior: " + e.toString());
        }
      }
      if (t.vencimiento) {
        try {
          agendarTareaEnCalendar(filaActual);
        } catch (e) {
          console.log("Error al reagendar en calendario: " + e.toString());
        }
      }
    }

  } else {
    // TAREA NUEVA
    idTareaActual = data.length > 1 ? Number(data[data.length - 1][1]) + 1 : 1;
    if (isNaN(idTareaActual)) idTareaActual = 1;

    const filaValores = [
      new Date(),
      idTareaActual,
      t.compania,
      t.tipoTarea,
      t.descripcion,
      t.vencimiento ? new Date(t.vencimiento + "T12:00:00") : "",
      t.estado,
      "",
      "",
      t.prioridad,
      t.adjunto,
      t.idCliente,
      usuarioActivo || "Sistema",
      t.responsable || "",
      t.ramo || ""
    ];

    hoja.appendRow(filaValores);
    
    if (t.vencimiento) {
      try {
        const nuevaFila = hoja.getLastRow();
        agendarTareaEnCalendar(nuevaFila);
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
 * CREA EL EVENTO EN EL CALENDARIO QUE CORRESPONDE AL RESPONSABLE
 * Recibe el NÚMERO DE FILA de la hoja TAREAS y lee todo desde ahí.
 */
function agendarTareaEnCalendar(numeroFilaReal) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName("TAREAS");
    const fila = hoja.getRange(numeroFilaReal, 1, 1, 15).getValues()[0];

    const idCliente = fila[11];          // Columna L
    const nombreTarea = fila[3];         // Columna D
    const fechaVencimientoRaw = fila[5]; // Columna F
    const responsable = (fila[13] || "").toString().trim(); // Columna N

    if (!fechaVencimientoRaw) {
      Logger.log(`Renglón ${numeroFilaReal} no tiene fecha de vencimiento, se omite.`);
      return;
    }

    const hojaClientes = ss.getSheetByName("CLIENTES");
    const datosClientes = hojaClientes.getDataRange().getValues();
    let nombre = "Cliente desconocido";
    let telefono = "No cargado";
    let email = "";

    for (let i = 1; i < datosClientes.length; i++) {
      if (datosClientes[i][0].toString() === idCliente.toString()) { 
        nombre = datosClientes[i][1];
        telefono = datosClientes[i][4];
        email = datosClientes[i][5];
        break;
      }
    }

    // Elegimos el calendario según el responsable
    const idCalendarioDestino = CALENDARIOS_POR_RESPONSABLE[responsable] || CALENDARIO_POR_DEFECTO;
    const calendario = CalendarApp.getCalendarById(idCalendarioDestino);

    if (!calendario) {
      Logger.log(`No se pudo abrir el calendario para "${responsable}" (${idCalendarioDestino}). Revisar permisos o ID.`);
      return;
    }

    const titulo = "Tarea: " + nombreTarea + " - " + nombre + (responsable ? " (" + responsable + ")" : "");
    
    let fechaEvento;
    if (fechaVencimientoRaw instanceof Date) {
      fechaEvento = new Date(fechaVencimientoRaw);
      fechaEvento.setHours(9, 0, 0, 0);
    } else {
      fechaEvento = new Date(fechaVencimientoRaw + "T09:00:00");
    }
    const finEvento = new Date(fechaEvento.getTime() + 60 * 60 * 1000);

    const descripcion = "ID Cliente: " + idCliente +
                        "\nNombre: " + nombre + 
                        "\nTarea: " + nombreTarea + 
                        "\nResponsable: " + (responsable || "Sin asignar") +
                        "\nTeléfono: " + telefono + 
                        "\nEmail: " + email;

    const evento = calendario.createEvent(titulo, fechaEvento, finEvento, {
      description: descripcion
    });

    evento.addEmailReminder(1440);

    hoja.getRange(numeroFilaReal, 9).setValue(evento.getId());
    
    Logger.log(`Evento creado para el renglón ${numeroFilaReal} en calendario de "${responsable || "por defecto"}"`);
    return evento.getId();
    
  } catch (e) {
    Logger.log("Error interno en agendarTareaEnCalendar (fila " + numeroFilaReal + "): " + e.toString());
  }
}

function obtenerResponsables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("LOGIN");
  if (!hoja) return [];
  const datos = hoja.getDataRange().getValues();
  return datos.slice(1).map(fila => fila[2]).filter(nombre => nombre);
}

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

function eliminarTareaEnServidor(idFila) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  
  const filaNumero = Number(idFila);
  
  if (filaNumero > 1 && filaNumero <= hoja.getLastRow()) {
    hoja.deleteRow(filaNumero);
    return { exito: true };
  } else {
    throw new Error("Número de fila inválido para eliminar.");
  }
}

/**
 * Función automatizada (trigger horario)
 * Recorre TAREAS y agenda las que no tengan ID de evento en columna I
 * y no estén Terminado.
 */
function sincronizarTareasConCalendario() {
  const TIEMPO_MAXIMO_MS = 4.5 * 60 * 1000;
  const inicio = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const valoresEstado = hoja.getRange(2, 7, ultimaFila - 1, 1).getValues();
  const valoresIds = hoja.getRange(2, 9, ultimaFila - 1, 1).getValues();

  let procesadas = 0;

  for (let i = 0; i < valoresIds.length; i++) {
    if (Date.now() - inicio > TIEMPO_MAXIMO_MS) {
      Logger.log(`Corte por tiempo. Procesadas ${procesadas} filas en esta corrida.`);
      break;
    }

    const idExistente = valoresIds[i][0];
    const estado = (valoresEstado[i][0] || "").toString().trim();
    const numeroFilaReal = i + 2;

    if (idExistente && idExistente.toString().trim() !== "") continue;
    if (estado === "Terminado") continue;

    try {
      agendarTareaEnCalendar(numeroFilaReal);
      procesadas++;
    } catch (error) {
      Logger.log(`Error al procesar el renglón ${numeroFilaReal}: ${error.toString()}`);
    }
  }

  Logger.log(`Sincronización finalizada. Total procesadas: ${procesadas}`);
}

/**
 * Función diaria: borra eventos de Calendar de tareas Terminado
 * y limpia la columna I.
 */
function borrarEventosDeTareasTerminadas() {
  const TIEMPO_MAXIMO_MS = 4.5 * 60 * 1000;
  const inicio = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 15).getValues();
  const clientes = obtenerListaClientes();

  // Revisamos el calendario por defecto + todos los calendarios de responsables
  const idsCalendarios = [CALENDARIO_POR_DEFECTO, ...Object.values(CALENDARIOS_POR_RESPONSABLE)];
  const calendariosUnicos = [...new Set(idsCalendarios)];
  const calendarios = calendariosUnicos.map(id => CalendarApp.getCalendarById(id)).filter(c => c);

  let borrados = 0;
  let noEncontrados = 0;

  for (let i = 0; i < datos.length; i++) {
    if (Date.now() - inicio > TIEMPO_MAXIMO_MS) {
      Logger.log(`Corte por tiempo. Borrados ${borrados} en esta corrida. Volvé a ejecutar para seguir con el resto.`);
      break;
    }

    const numeroFilaReal = i + 2;
    const estado = (datos[i][6] || "").toString().trim();       // G
    const idEvento = (datos[i][8] || "").toString().trim();      // I
    const idCliente = datos[i][11];                              // L
    const tipoTarea = datos[i][3];                                // D
    const responsable = (datos[i][13] || "").toString().trim();  // N
    const vencimientoRaw = datos[i][5];                           // F

    if (estado !== "Terminado") continue;

    let eventoBorrado = false;

    // 1. Intento directo con el ID guardado, buscando en cualquier calendario accesible
    if (idEvento !== "") {
      try {
        const evento = CalendarApp.getEventById(idEvento);
        if (evento) {
          evento.deleteEvent();
          eventoBorrado = true;
        }
      } catch (error) {
        Logger.log(`Error al borrar por ID en renglón ${numeroFilaReal}: ${error.toString()}`);
      }
      hoja.getRange(numeroFilaReal, 9).setValue("");
    }

    // 2. Fallback: si no había ID o no se encontró el evento, buscamos por fecha + título
    if (!eventoBorrado && vencimientoRaw instanceof Date && !isNaN(vencimientoRaw)) {
      const cliente = clientes.find(c => c.id == idCliente);
      const nombreCliente = cliente ? cliente.nombrePuro : null;

      if (nombreCliente) {
        const tituloConResponsable = "Tarea: " + tipoTarea + " - " + nombreCliente + (responsable ? " (" + responsable + ")" : "");
        const tituloSinResponsable = "Tarea: " + tipoTarea + " - " + nombreCliente;

        const inicioDia = new Date(vencimientoRaw);
        inicioDia.setHours(0, 0, 0, 0);
        const finDia = new Date(vencimientoRaw);
        finDia.setHours(23, 59, 59, 999);

        for (const cal of calendarios) {
          const eventosDelDia = cal.getEvents(inicioDia, finDia);
          eventosDelDia.forEach(ev => {
            const titulo = ev.getTitle();
            if (titulo === tituloConResponsable || titulo === tituloSinResponsable) {
              ev.deleteEvent();
              eventoBorrado = true;
            }
          });
        }
      }
    }

    if (eventoBorrado) borrados++;
    else noEncontrados++;
  }

  Logger.log(`Finalizado. Total eventos borrados: ${borrados} | No encontrados/sin match: ${noEncontrados}`);
}

/**
 * MIGRACIÓN (ejecutar manualmente 1 sola vez, o varias veces si corta por tiempo)
 * Mueve al calendario correcto las tareas de responsables que ya tienen evento
 * creado en el calendario por defecto. Con el mapa actual, solo migra las de Matias.
 */
function migrarEventosARepsonsables() {
  const TIEMPO_MAXIMO_MS = 4.5 * 60 * 1000;
  const inicio = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("TAREAS");
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const datos = hoja.getRange(2, 1, ultimaFila - 1, 15).getValues();
  const calViejo = CalendarApp.getCalendarById(CALENDARIO_POR_DEFECTO);

  let migradas = 0;

  for (let i = 0; i < datos.length; i++) {
    if (Date.now() - inicio > TIEMPO_MAXIMO_MS) {
      Logger.log(`Corte por tiempo. Migradas ${migradas} en esta corrida. Volvé a ejecutar para seguir.`);
      break;
    }

    const numeroFilaReal = i + 2;
    const estado = (datos[i][6] || "").toString().trim();       // G
    const idEventoViejo = (datos[i][8] || "").toString().trim(); // I
    const responsable = (datos[i][13] || "").toString().trim();  // N

    if (estado === "Terminado") continue;
    if (idEventoViejo === "") continue;
    if (!CALENDARIOS_POR_RESPONSABLE[responsable]) continue; // no tiene calendario propio, se queda como está

    try {
      const eventoViejo = calViejo ? calViejo.getEventById(idEventoViejo) : null;
      if (eventoViejo) eventoViejo.deleteEvent();

      agendarTareaEnCalendar(numeroFilaReal);
      migradas++;
      Logger.log(`Renglón ${numeroFilaReal} migrado al calendario de "${responsable}"`);
    } catch (error) {
      Logger.log(`Error migrando renglón ${numeroFilaReal}: ${error.toString()}`);
    }
  }

  Logger.log(`Migración finalizada. Total migradas: ${migradas}`);
}
// --- GESTIÓN DE RENOVACIONES ---

/**
 * Devuelve todas las renovaciones cargadas en la hoja RENOVACIONES.
 * Estructura esperada de la hoja (fila 1 = encabezados):
 * A: ID | B: FechaCreacion | C: ID_Cliente | D: FechaRenovacion ("DD/MM") | E: Ramo |
 * F: Comentario | G: Usuario | H: Frecuencia | I: UltimaTareaGenerada | J: Compania
 *
 * Como la fecha no tiene año (se repite todos los años), acá calculamos la
 * "próxima ocurrencia" usando el año actual del servidor, para poder
 * ordenar/filtrar (Vencidas / Vence en 1 Semana) igual que con las tareas.
 */
function obtenerRenovaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("RENOVACIONES");
  if (!hoja) return [];

  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  const clientes = obtenerListaClientes();
  const anioActual = new Date().getFullYear();

  return data.slice(1).map((r, index) => {
    const idCliente = r[2];                                    // Columna C
    const ramo = r[4] || "";                                    // Columna E
    const comentario = r[5] || "";                              // Columna F
    const frecuencia = r[7] || "";                              // Columna H
    const ultimaTareaGeneradaRaw = r[8];                        // Columna I
    const compania = r[9] || "";                                // Columna J

    if (!idCliente) return null; // fila vacía, la filtramos después

    const c = clientes.find(x => x.id == idCliente);

    // La columna D puede venir como texto "DD/MM" (caso normal) o, si Sheets
    // la autoconvirtió a fecha real en algún momento, como objeto Date.
    // Contemplamos los dos casos para no perder el dato.
    let dia = null, mes = null, fechaRenovacionRaw = "";
    const celdaFecha = r[3];
    if (celdaFecha instanceof Date) {
      dia = celdaFecha.getDate();
      mes = celdaFecha.getMonth() + 1;
      fechaRenovacionRaw = (dia < 10 ? "0" + dia : dia) + "/" + (mes < 10 ? "0" + mes : mes);
    } else {
      fechaRenovacionRaw = (celdaFecha || "").toString().trim();
      if (fechaRenovacionRaw.indexOf("/") !== -1) {
        const partes = fechaRenovacionRaw.split("/");
        dia = parseInt(partes[0], 10);
        mes = parseInt(partes[1], 10);
      }
    }

    let proximaISO = "";
    let proximaFormat = "";
    if (dia && mes && !isNaN(dia) && !isNaN(mes)) {
      const fechaEsteAnio = new Date(anioActual, mes - 1, dia);
      proximaISO = Utilities.formatDate(fechaEsteAnio, "GMT-3", "yyyy-MM-dd");
      proximaFormat = Utilities.formatDate(fechaEsteAnio, "GMT-3", "dd/MM/yyyy");
    }

    let fCreacionRaw = r[1]; // Columna B
    let fCreacionFormat = (fCreacionRaw instanceof Date) ? Utilities.formatDate(fCreacionRaw, "GMT-3", "dd/MM/yy") : "-";

    const ultimaTareaGenerada = (ultimaTareaGeneradaRaw instanceof Date)
      ? Utilities.formatDate(ultimaTareaGeneradaRaw, "GMT-3", "dd/MM/yyyy")
      : "";

    return {
      id_fila: index + 2,
      idCliente: idCliente,
      clienteNombre: c ? c.nombre : "Sin asignar",
      creacion: fCreacionFormat,
      fechaRenovacion: fechaRenovacionRaw, // "DD/MM" tal cual está guardado
      vencimiento: proximaISO,             // próxima ocurrencia (recalcula sola cada año)
      vencimientoFormat: proximaFormat,
      ramo: ramo,
      frecuencia: frecuencia,
      compania: compania,
      ultimaTareaGenerada: ultimaTareaGenerada,
      comentario: comentario
    };
  }).filter(r => r !== null);
}

/**
 * Devuelve solo las renovaciones de un cliente puntual (para la pantalla de Clientes).
 */
function obtenerRenovacionesPorCliente(idCliente) {
  return obtenerRenovaciones().filter(r => r.idCliente == idCliente);
}

// Cuántos meses hay entre una renovación y la siguiente, según la frecuencia elegida
const INTERVALO_MESES_POR_FRECUENCIA = {
  "Mensual": 1,
  "Trimestral": 3,
  "Cuatrimestral": 4,
  "Semestral": 6,
  "Anual": 12
};

/**
 * Calcula las próximas "cantidad" ocurrencias de una fecha día/mes recurrente,
 * espaciadas cada "intervaloMeses" meses, empezando desde "desde" (o desde hoy si no se pasa).
 * La primera ocurrencia siempre es >= "desde".
 */
function calcularProximasOcurrencias(dia, mes, intervaloMeses, cantidad, desde) {
  const inicio = desde ? new Date(desde) : new Date();
  inicio.setHours(0, 0, 0, 0);

  let fecha = new Date(inicio.getFullYear(), mes - 1, dia);
  while (fecha < inicio) {
    fecha = new Date(fecha.getFullYear(), fecha.getMonth() + intervaloMeses, fecha.getDate());
  }

  const ocurrencias = [];
  let actual = fecha;
  for (let i = 0; i < cantidad; i++) {
    ocurrencias.push(new Date(actual));
    actual = new Date(actual.getFullYear(), actual.getMonth() + intervaloMeses, actual.getDate());
  }
  return ocurrencias;
}

/**
 * Crea en la hoja TAREAS una tarea de tipo "Renovacion" por cada fecha de la lista,
 * agendando también su evento de Calendar (igual que una tarea creada a mano).
 * Devuelve la cantidad de tareas creadas.
 */
function generarTareasRenovacion(idCliente, compania, ramo, comentario, fechasOcurrencias, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) return 0;

  const data = hojaTareas.getDataRange().getValues();
  let idTareaActual = data.length > 1 ? Number(data[data.length - 1][1]) : 0;
  if (isNaN(idTareaActual)) idTareaActual = 0;

  let creadas = 0;
  fechasOcurrencias.forEach(fecha => {
    idTareaActual++;
    const fechaVenc = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), 12, 0, 0);

    const filaValores = [
      new Date(),          // A: Fecha Creación
      idTareaActual,       // B: ID Tarea
      compania || "",      // C
      "Renovacion",        // D: Tipo de Tarea
      comentario || "",    // E: Descripción
      fechaVenc,           // F: Vencimiento
      "Sin leer",          // G: Estado
      "",                  // H
      "",                  // I
      "Normal",            // J: Prioridad
      "",                  // K: Adjunto
      idCliente,           // L
      usuarioActivo || "Sistema", // M
      "",                  // N: Responsable
      ramo || ""           // O: Ramo
    ];

    hojaTareas.appendRow(filaValores);
    creadas++;

    try {
      const nuevaFila = hojaTareas.getLastRow();
      agendarTareaEnCalendar(nuevaFila);
    } catch (e) {
      Logger.log("Error al agendar tarea de renovación en calendario: " + e.toString());
    }
  });

  return creadas;
}

/**
 * Agrega una nueva renovación a la hoja RENOVACIONES y crea de una vez las próximas
 * 10 tareas de tipo "Renovacion" en TAREAS (con sus eventos de Calendar), espaciadas
 * según la frecuencia elegida.
 * datos = { idCliente, fechaRenovacion ("DD/MM"), frecuencia, compania, ramo, comentario }
 */
function guardarRenovacion(datos, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("RENOVACIONES");
  if (!hoja) throw new Error("No se encontró la hoja RENOVACIONES");

  const data = hoja.getDataRange().getValues();
  const idNuevo = data.length > 1 ? Number(data[data.length - 1][0]) + 1 : 1;

  const partes = datos.fechaRenovacion.split("/");
  const dia = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10);
  const intervaloMeses = INTERVALO_MESES_POR_FRECUENCIA[datos.frecuencia] || 12;

  const ocurrencias = calcularProximasOcurrencias(dia, mes, intervaloMeses, 10);
  const ultimaOcurrencia = ocurrencias[ocurrencias.length - 1];

  hoja.appendRow([
    idNuevo,
    new Date(),
    datos.idCliente,
    "",
    datos.ramo || "",
    datos.comentario || "",
    usuarioActivo || "Sistema",
    datos.frecuencia || "",
    ultimaOcurrencia,
    datos.compania || ""
  ]);

  // Forzamos la columna D (FechaRenovacion) a texto plano ANTES de escribir el valor,
  // para que Sheets no la autoconvierta a una fecha real (nos rompía el "DD/MM").
  const filaNueva = hoja.getLastRow();
  const celdaFecha = hoja.getRange(filaNueva, 4);
  celdaFecha.setNumberFormat("@");
  celdaFecha.setValue(datos.fechaRenovacion);

  const creadas = generarTareasRenovacion(datos.idCliente, datos.compania, datos.ramo, datos.comentario, ocurrencias, usuarioActivo);

  return {
    exito: true,
    tareasCreadas: creadas,
    hasta: Utilities.formatDate(ultimaOcurrencia, "GMT-3", "dd/MM/yyyy")
  };
}

/**
 * Genera manualmente las próximas 10 tareas de una renovación ya existente,
 * continuando desde la última tarea que se le había generado (columna I).
 * Se llama con el id_fila de la renovación (fila real en la hoja RENOVACIONES).
 */
function regenerarTareasRenovacion(idFilaRenovacion) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("RENOVACIONES");
  if (!hoja) throw new Error("No se encontró la hoja RENOVACIONES");

  const fila = Number(idFilaRenovacion);
  const datosFila = hoja.getRange(fila, 1, 1, 10).getValues()[0];

  const idCliente = datosFila[2];
  const celdaFecha = datosFila[3];
  const ramo = datosFila[4];
  const comentario = datosFila[5];
  const frecuencia = datosFila[7];
  const ultimaGenerada = datosFila[8];
  const compania = datosFila[9];

  let dia, mes;
  if (celdaFecha instanceof Date) {
    dia = celdaFecha.getDate();
    mes = celdaFecha.getMonth() + 1;
  } else {
    const partes = (celdaFecha || "").toString().split("/");
    dia = parseInt(partes[0], 10);
    mes = parseInt(partes[1], 10);
  }

  const intervaloMeses = INTERVALO_MESES_POR_FRECUENCIA[frecuencia] || 12;

  // Arrancamos el día después de la última tarea que ya se había generado
  let desde;
  if (ultimaGenerada instanceof Date) {
    desde = new Date(ultimaGenerada);
    desde.setDate(desde.getDate() + 1);
  } else {
    desde = new Date();
  }

  const ocurrencias = calcularProximasOcurrencias(dia, mes, intervaloMeses, 10, desde);
  const ultimaOcurrencia = ocurrencias[ocurrencias.length - 1];

  const creadas = generarTareasRenovacion(idCliente, compania, ramo, comentario, ocurrencias, "Sistema");

  hoja.getRange(fila, 9).setValue(ultimaOcurrencia); // actualiza UltimaTareaGenerada (columna I)

  return {
    exito: true,
    tareasCreadas: creadas,
    hasta: Utilities.formatDate(ultimaOcurrencia, "GMT-3", "dd/MM/yyyy")
  };
}

/**
 * EJECUTAR MANUALMENTE 1 SOLA VEZ si tenías renovaciones cargadas antes de este fix,
 * cuya columna D (FechaRenovacion) haya quedado autoconvertida a una fecha real en
 * vez de quedar como texto "DD/MM". Recorre la hoja y las corrige.
 */
function repararFechasRenovaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("RENOVACIONES");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rango = hoja.getRange(2, 4, ultimaFila - 1, 1); // columna D
  const valores = rango.getValues();
  let corregidas = 0;

  for (let i = 0; i < valores.length; i++) {
    const valor = valores[i][0];
    if (valor instanceof Date) {
      const dia = valor.getDate();
      const mes = valor.getMonth() + 1;
      const texto = (dia < 10 ? "0" + dia : dia) + "/" + (mes < 10 ? "0" + mes : mes);
      const celda = hoja.getRange(i + 2, 4);
      celda.setNumberFormat("@");
      celda.setValue(texto);
      corregidas++;
    }
  }

  Logger.log(`Fechas corregidas: ${corregidas}`);
}

function eliminarRenovacionEnServidor(idFila) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("RENOVACIONES");
  if (!hoja) throw new Error("No se encontró la hoja RENOVACIONES");

  const filaNumero = Number(idFila);

  if (filaNumero > 1 && filaNumero <= hoja.getLastRow()) {
    hoja.deleteRow(filaNumero);
    return { exito: true };
  } else {
    throw new Error("Número de fila inválido para eliminar.");
  }
}