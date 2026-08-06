// --- GESTIÓN DE CLIENTES ---
// (No confundir con Clientes.gs, que es el script de sincronización automática
// desde LIQUIDACIONES — este archivo es la parte de la webapp del CRM.)

// Etiqueta de compañías [RIV / PS / FP] que se graba directo en la columna B
// (Nombre) de CLIENTES, para que la hoja sea la fuente de verdad y la webapp
// no tenga que calcular ni agregar nada por su cuenta.
function calcularEtiquetaCompanias(rivadavia, provincia, fedPatronal) {
  let etiquetas = [];
  if (rivadavia) etiquetas.push("RIV");
  if (provincia) etiquetas.push("PS");
  if (fedPatronal) etiquetas.push("FP");
  return etiquetas.length > 0 ? " [" + etiquetas.join(" / ") + "]" : "";
}

// Saca cualquier etiqueta "[...]" que ya esté pegada a un nombre, para poder
// recalcularla de cero sin ir acumulando corchetes viejos.
function limpiarEtiquetaNombre(nombre) {
  return (nombre || "").toString().replace(/ \[[^\]]*\]/g, "").trim();
}

function obtenerListaClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  const data = hoja.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Armamos un mapa Nombre -> "Renovación Combinado Familiar" leyendo la hoja
  // FORMULARIO CF: columna O = nombre para vincular con CLIENTES columna B,
  // columna P = el dato de solo lectura que se muestra en la ficha del cliente.
  const mapaRenovacionCF = {};
  const hojaCF = ss.getSheetByName("FORMULARIO CF");
  if (hojaCF) {
    const datosCF = hojaCF.getDataRange().getValues();
    for (let i = 1; i < datosCF.length; i++) {
      const nombreCF = limpiarEtiquetaNombre(datosCF[i][14]).toUpperCase(); // Columna O, sin el "[RIV/PS/FP]"
      if (nombreCF) {
        mapaRenovacionCF[nombreCF] = datosCF[i][15]; // Columna P
      }
    }
  }
  
  return data.slice(1).map(r => {
    const nombreClienteLimpio = limpiarEtiquetaNombre(r[1]).toUpperCase();
    return {
      id: r[0],
      nombre: r[1], // tal cual está en la hoja (ya incluye "[RIV / PS / FP]" si corresponde)
      dni: r[2],
      domicilio: r[3],
      telefono: r[4],
      email: r[5],
      rivadavia: r[6],
      provincia: r[7],
      fedPatronal: r[8],
      relacionados: r[9],
      observaciones: r[10],
      renovacionCombinadoFamiliar: mapaRenovacionCF[nombreClienteLimpio] || ""
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

  // El nombre siempre se guarda con la etiqueta de compañías recalculada de cero,
  // así la hoja CLIENTES queda como fuente de verdad (y la webapp solo muestra
  // lo que ya está ahí, sin agregar nada por su cuenta).
  const nombreLimpio = limpiarEtiquetaNombre(datos.nombre.toUpperCase());
  const etiqueta = calcularEtiquetaCompanias(datos.rivadavia, datos.provincia, datos.fedPatronal);
  const nombreConEtiqueta = nombreLimpio + etiqueta;

  hoja.getRange(filaIndex, 2, 1, 10).setValues([[
    nombreConEtiqueta, 
    datos.dni, 
    datos.domicilio, 
    datos.telefono, 
    datos.email, 
    datos.rivadavia,
    datos.provincia,
    datos.fedPatronal,
    relacionadosJSON,
    datos.observaciones || ""
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

// --- MANTENIMIENTO DE NOMBRES/ETIQUETAS ---

/**
 * EJECUTAR MANUALMENTE 1 SOLA VEZ para arreglar los nombres de clientes que
 * quedaron con "[RIV]", "[PS]" o "[FP]" grabados como texto literal en la
 * columna B (bug de cargarDatosClienteEdicion, ya corregido). Saca cualquier
 * cantidad de esas etiquetas que haya quedado pegada al nombre, incluidas
 * las duplicadas (ej: "CARDOSO GUSTAVO [FP] [FP]" -> "CARDOSO GUSTAVO").
 */
function limpiarNombresClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rango = hoja.getRange(2, 2, ultimaFila - 1, 1); // Columna B: Nombre
  const valores = rango.getValues();
  const regexEtiquetas = / \[[^\]]*\]/g;
  let corregidos = 0;

  for (let i = 0; i < valores.length; i++) {
    const nombreOriginal = (valores[i][0] || "").toString();
    const nombreLimpio = nombreOriginal.replace(regexEtiquetas, "").trim();
    if (nombreLimpio !== nombreOriginal) {
      valores[i][0] = nombreLimpio;
      corregidos++;
    }
  }

  rango.setValues(valores);
  Logger.log(`Nombres corregidos: ${corregidos}`);
}

/**
 * Trigger simple: se dispara solo cada vez que se edita CUALQUIER celda del
 * spreadsheet. Si la edición fue en las columnas G, H o I (Rivadavia,
 * Provincia, FedPatronal) de la hoja CLIENTES, recalcula la etiqueta
 * "[RIV / PS / FP]" y la actualiza en la columna B (Nombre) de esa fila.
 * Así, aunque cargues o edites números directo en el Sheet (sin pasar por
 * la webapp), el nombre se mantiene sincronizado solo.
 */
function onEdit(e) {
  try {
    const hoja = e.range.getSheet();
    if (hoja.getName() !== "CLIENTES") return;

    const fila = e.range.getRow();
    if (fila === 1) return; // encabezado

    const columnaInicio = e.range.getColumn();
    const columnaFin = e.range.getLastColumn();
    // G=7 (Rivadavia), H=8 (Provincia), I=9 (FedPatronal)
    const tocaColumnasDeCompania = columnaInicio <= 9 && columnaFin >= 7;
    if (!tocaColumnasDeCompania) return;

    actualizarEtiquetaClienteEnFila(hoja, fila);
  } catch (err) {
    Logger.log("Error en onEdit (CLIENTES): " + err.toString());
  }
}

function actualizarEtiquetaClienteEnFila(hoja, fila) {
  const valores = hoja.getRange(fila, 2, 1, 8).getValues()[0]; // B..I
  const nombreActual = valores[0];
  const rivadavia = valores[5];
  const provincia = valores[6];
  const fedPatronal = valores[7];

  if (!nombreActual) return; // fila vacía, no tocamos nada

  const nombreLimpio = limpiarEtiquetaNombre(nombreActual);
  const etiqueta = calcularEtiquetaCompanias(rivadavia, provincia, fedPatronal);
  hoja.getRange(fila, 2).setValue(nombreLimpio + etiqueta);
}

/**
 * EJECUTAR MANUALMENTE 1 SOLA VEZ para poner al día TODAS las etiquetas de
 * golpe (por ejemplo, la primera vez que se activa este sistema, o si se
 * importaron/editaron muchas filas de una y el trigger no llegó a correr
 * para cada una). Recalcula la etiqueta de cada cliente según sus columnas
 * G/H/I actuales y la deja grabada en la columna B.
 */
function regenerarEtiquetasClientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("CLIENTES");
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return;

  const rango = hoja.getRange(2, 2, ultimaFila - 1, 8); // B..I
  const valores = rango.getValues();

  const nombresNuevos = valores.map(fila => {
    const nombreActual = fila[0];
    if (!nombreActual) return [nombreActual];
    const nombreLimpio = limpiarEtiquetaNombre(nombreActual);
    const etiqueta = calcularEtiquetaCompanias(fila[5], fila[6], fila[7]);
    return [nombreLimpio + etiqueta];
  });

  hoja.getRange(2, 2, nombresNuevos.length, 1).setValues(nombresNuevos);
  Logger.log(`Nombres actualizados: ${nombresNuevos.length}`);
}