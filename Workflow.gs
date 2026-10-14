/**
 * ===============================================
 * ARCHIVO: Workflow.gs
 * Gestión del Flujo de Trabajo (Marcar como Hecho)
 * Mapeo Dinámico por Encabezados (Fila 1)
 * ===============================================
 */

function marcarCortesComoHechos(hojaDatos, todosLosIdsParaMarcar) { 
  const idsAMarcar = new Set(todosLosIdsParaMarcar);
  const ultimaFila = hojaDatos.getLastRow();
  
  if (ultimaFila < FILA_INICIO_DATOS) return;

  const colMap = getHeaderMap(hojaDatos, FILA_CABECERA);
  const colId = colMap.getCol("Id_Ventana") || colMap.getCol("ID");
  const colHecho = colMap.getCol("Estado_Ventana") || colMap.getCol("Calculado") || colMap.getCol("Hecha");

  if (!colId || !colHecho) {
    console.error("[Workflow] No se pudieron localizar las columnas 'Id_Ventana' o 'Estado_Ventana'.");
    return;
  }

  const numFilas = ultimaFila - FILA_INICIO_DATOS + 1;
  const rangoIDs = hojaDatos.getRange(FILA_INICIO_DATOS, colId, numFilas, 1).getValues();
  const rangoHecho = hojaDatos.getRange(FILA_INICIO_DATOS, colHecho, numFilas, 1);
  const valoresHecho = rangoHecho.getValues();
  
  let cambios = false;
  let trabajosMarcados = 0;

  for (let i = 0; i < rangoIDs.length; i++) {
    const idFila = rangoIDs[i][0];
    if (idsAMarcar.has(idFila) && (valoresHecho[i][0] !== true)) {
      valoresHecho[i][0] = true;
      trabajosMarcados++;
      cambios = true;
    }
  }

  if (cambios) {
    rangoHecho.setValues(valoresHecho);
    console.log(`[Workflow] ${trabajosMarcados} filas de trabajo marcadas como HECHAS.`);
  } else {
    console.log(`[Workflow] No se encontraron nuevos trabajos para marcar.`);
  }
}