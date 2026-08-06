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
 * agendando también su evento de Calendar. Devuelve la cantidad de tareas creadas.
 *
 * OPTIMIZADA: antes esta función releía TODA la hoja CLIENTES y hacía un appendRow +
 * una relectura de la fila por cada una de las 10 tareas (a través de
 * agendarTareaEnCalendar), lo que la hacía muy lenta. Ahora el dato del cliente se
 * busca una única vez y las filas de TAREAS se escriben todas juntas en un solo batch.
 * Lo único que sigue siendo 1 llamada por tarea es la creación del evento de Calendar,
 * porque la API de Calendar no permite crear varios eventos en una sola llamada.
 */
function generarTareasRenovacion(idCliente, compania, ramo, comentario, fechasOcurrencias, usuarioActivo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTareas = ss.getSheetByName("TAREAS");
  if (!hojaTareas) return 0;

  const ultimaFilaActual = hojaTareas.getLastRow();
  let idTareaActual = 0;
  if (ultimaFilaActual > 1) {
    idTareaActual = Number(hojaTareas.getRange(ultimaFilaActual, 2).getValue()) || 0;
  }

  // Buscamos los datos del cliente UNA sola vez (antes se releía toda la hoja CLIENTES por cada tarea)
  const hojaClientes = ss.getSheetByName("CLIENTES");
  const datosClientes = hojaClientes.getDataRange().getValues();
  let nombreCliente = "Cliente desconocido";
  let telefono = "No cargado";
  let email = "";
  for (let i = 1; i < datosClientes.length; i++) {
    if (datosClientes[i][0].toString() === idCliente.toString()) {
      nombreCliente = limpiarEtiquetaNombre(datosClientes[i][1]);
      telefono = datosClientes[i][4];
      email = datosClientes[i][5];
      break;
    }
  }

  // Armamos todas las filas en memoria y las escribimos en una sola operación
  const filasNuevas = fechasOcurrencias.map(fecha => {
    idTareaActual++;
    const fechaVenc = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate(), 12, 0, 0);
    return [
      new Date(),          // A: Fecha Creación
      idTareaActual,       // B: ID Tarea
      compania || "",      // C
      "Renovacion",        // D: Tipo de Tarea
      comentario || "",    // E: Descripción
      fechaVenc,           // F: Vencimiento
      "Sin leer",          // G: Estado
      "",                  // H
      "",                  // I: ID evento de Calendar (se completa después)
      "Normal",            // J: Prioridad
      "",                  // K: Adjunto
      idCliente,           // L
      usuarioActivo || "Sistema", // M
      "",                  // N: Responsable
      ramo || ""           // O: Ramo
    ];
  });

  const filaInicio = ultimaFilaActual + 1;
  hojaTareas.getRange(filaInicio, 1, filasNuevas.length, 15).setValues(filasNuevas);

  // Creamos los eventos de Calendar. Esta parte sigue siendo 1 llamada por evento
  // (la API de Calendar no admite creación en lote), pero ya no repetimos lecturas.
  const calendario = CalendarApp.getCalendarById(CALENDARIO_POR_DEFECTO);
  const idsEventos = filasNuevas.map(fila => {
    let idEvento = "";
    if (calendario) {
      try {
        const fechaEvento = fila[5];
        const finEvento = new Date(fechaEvento.getTime() + 60 * 60 * 1000);
        const titulo = "Tarea: Renovacion - " + nombreCliente;
        const descripcion = "ID Cliente: " + idCliente +
                            "\nNombre: " + nombreCliente +
                            "\nTarea: Renovacion" +
                            "\nTeléfono: " + telefono +
                            "\nEmail: " + email;
        const evento = calendario.createEvent(titulo, fechaEvento, finEvento, { description: descripcion });
        evento.addEmailReminder(1440);
        idEvento = evento.getId();
      } catch (e) {
        Logger.log("Error al agendar tarea de renovación en calendario: " + e.toString());
      }
    }
    return [idEvento];
  });

  // Guardamos todos los IDs de evento (columna I) en un solo batch
  hojaTareas.getRange(filaInicio, 9, idsEventos.length, 1).setValues(idsEventos);

  return filasNuevas.length;
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