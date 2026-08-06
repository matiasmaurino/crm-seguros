// --- MAPEO DE RESPONSABLE -> CALENDARIO ---
const CALENDARIOS_POR_RESPONSABLE = {
  "Matias": "0q14of40e34nhec1lnjh2kami8@group.calendar.google.com"
};
const CALENDARIO_POR_DEFECTO = "matiasmaurino@gmail.com"; // Mariana, Mauro, Pilar y sin asignar caen acá

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
        nombre = limpiarEtiquetaNombre(datosClientes[i][1]);
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
      const nombreCliente = cliente ? limpiarEtiquetaNombre(cliente.nombre) : null;

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