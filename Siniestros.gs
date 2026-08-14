// --- MÓDULO SINIESTROS ---
// Reutiliza obtenerTareas() (con el mismo cruce de cliente, ramo, etc.) y
// simplemente filtra las que son de Tipo de Tarea "Siniestro" y todavía no
// están Terminadas — esta pantalla es un recorte de Historial de Tareas
// para lo que sigue pendiente, no un archivo histórico.

function obtenerSiniestros() {
  const todasLasTareas = obtenerTareas();
  return todasLasTareas.filter(t =>
    (t.tipoTarea || "").toString().trim().toLowerCase() === "siniestro" &&
    (t.estado || "").toString().trim().toLowerCase() !== "terminado"
  );
}