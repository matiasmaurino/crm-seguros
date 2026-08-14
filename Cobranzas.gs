// --- MÓDULO COBRANZAS ---
// Reutiliza obtenerTareas() (mismo cruce de cliente, ramo, etc.) y filtra
// las que son de Tipo de Tarea "Cobranzas" y todavía no están Terminadas.

function obtenerCobranzas() {
  const todasLasTareas = obtenerTareas();
  return todasLasTareas.filter(t =>
    (t.tipoTarea || "").toString().trim().toLowerCase() === "cobranzas" &&
    (t.estado || "").toString().trim().toLowerCase() !== "terminado"
  );
}