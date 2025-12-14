/**
 * ===============================================
 * ARCHIVO: Workflow.gs
 * Gestión del Flujo de Trabajo (Marcar como Hecho)
 * ===============================================
 */

function marcarCortesComoHechos(hojaDatos, todosLosIdsParaMarcar) { 
  const idsAMarcar = new Set(todosLosIdsParaMarcar);
  const ultimaFila = hojaDatos.getLastRow();
  
  if (ultimaFila < FILA_INICIO_DATOS) return;

  const rangoIDs = hojaDatos.getRange(COLUMNA_ID_LETRA + FILA_INICIO_DATOS + ":" + COLUMNA_ID_LETRA + ultimaFila).getValues();
  const rangoHecho = hojaDatos.getRange(COLUMNA_HECHO_LETRA + FILA_INICIO_DATOS + ":" + COLUMNA_HECHO_LETRA + ultimaFila);  const valoresHecho = rangoHecho.getValues();
  
  let cambios = false;
  let trabajosMarcados = 0;

  for (let i = 0; i < rangoIDs.length; i++) {
    const idFila = rangoIDs[i][0];
      if (idsAMarcar.has(idFila) && (valoresHecho[i][0] !== true)) {
      valoresHecho[i][0] = true; // Marcamos el check
      cambios = true;
   }
  }

  if (cambios) {
    rangoHecho.setValues(valoresHecho);
    console.log(`[Marcado] ${trabajosMarcados} filas de trabajo marcadas como HECHAS.`);
  } else {
    console.log(`[Marcado] No se encontraron nuevos trabajos para marcar.`);
  }
}