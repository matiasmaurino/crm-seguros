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