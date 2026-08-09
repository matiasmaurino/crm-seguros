// --- MÓDULO SINIESTROS ---
// Reutiliza obtenerTareas() (con el mismo cruce de cliente, ramo, etc.) y
// simplemente filtra las que son de Tipo de Tarea "Siniestro". No duplica
// lógica, así que cualquier fix futuro en obtenerTareas() se hereda solo acá.

function obtenerSiniestros() {
  const todasLasTareas = obtenerTareas();
  return todasLasTareas.filter(t => (t.tipoTarea || "").toString().trim().toLowerCase() === "siniestro");
}