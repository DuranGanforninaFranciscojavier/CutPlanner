// ==============================================================================
// SECCIÓN: GENERADOR DE HOJA DE IMPRESIÓN COMBINADA (CORREGIDO)
// ==============================================================================

/**
 * Crea una hoja de impresión combinada.
 * 1. Filtra trabajos no calculados en 'Corredera_GP-60'.
 * 2. Tabla Superior: CLONA formato y valores de columnas específicas.
 * 3. Tabla Inferior: CLONA formato, valores y anchos de 'Vista_Tecnica...'.
 */
function generarImpresionCombinada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const nombreHojaOrigen = "Corredera_GP-60";
  const nombreHojaTecnica = "Vista_Tecnica Corredera_GP-60";
  const nombreHojaDestino = "Impresion_Combinada";
  
  const colIdIndex = 1;    // Col A en origen
  const colCheckIndex = 9; // Col I en origen
  const filaCabecera = 1;

  // Columnas para la tabla superior (Indices base 1)
  const colsVistaBase = [1, 2, 3, 5, 17, 18, 19, 20, 21, 22];

  // --- 1. SETUP Y VALIDACIONES ---
  const hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  if (!hojaOrigen) { SpreadsheetApp.getUi().alert(`Error: Falta la hoja "${nombreHojaOrigen}".`); return; }
  
  const hojaTecnica = ss.getSheetByName(nombreHojaTecnica);
  if (!hojaTecnica) { SpreadsheetApp.getUi().alert(`Error: Falta la hoja "${nombreHojaTecnica}".`); return; }

  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (hojaDestino) ss.deleteSheet(hojaDestino);
  hojaDestino = ss.insertSheet(nombreHojaDestino, ss.getNumSheets());


  // --- 2. FILTRADO Y MEMORIZACIÓN DE FILAS ---
  const ultimaFila = hojaOrigen.getLastRow();
  if (ultimaFila < 2) { SpreadsheetApp.getUi().alert("No hay datos para filtrar."); ss.deleteSheet(hojaDestino); return; }
  
  const rangoRevision = hojaOrigen.getRange(2, 1, ultimaFila - 1, colCheckIndex).getValues();
  const indicesFilasAImprimir = [];

  for (let i = 0; i < rangoRevision.length; i++) {
    const valorID = rangoRevision[i][colIdIndex - 1];
    const valorCheck = rangoRevision[i][colCheckIndex - 1];
    // Ajusta esta lógica según tus checks reales (true/false/string)
    if (valorID !== "" && valorID !== null && valorCheck !== true) {
      indicesFilasAImprimir.push(i + 2); // Guardamos índice real de la fila (base 1)
    }
  }

  if (indicesFilasAImprimir.length === 0) {
    SpreadsheetApp.getUi().alert("No hay trabajos pendientes para imprimir.");
    ss.deleteSheet(hojaDestino); return;
  }


  // --- 3. CONSTRUCCIÓN TABLA SUPERIOR (CLONADO CON FORMATO) ---
  // Copiamos la cabecera (Fila 1)
  copiarCeldasDiscontinuas(hojaOrigen, hojaDestino, filaCabecera, 1, colsVistaBase);

  // Copiamos las filas de datos filtradas
  for (let i = 0; i < indicesFilasAImprimir.length; i++) {
    // Fila destino: empieza en 2 (después de cabecera) + índice
    const filaDest = 2 + i; 
    copiarCeldasDiscontinuas(hojaOrigen, hojaDestino, indicesFilasAImprimir[i], filaDest, colsVistaBase);
  }
  
  // Opcional: Aplicar bordes exteriores a la tabla superior para limpieza visual
  const filasTablaSup = indicesFilasAImprimir.length + 1;
  hojaDestino.getRange(1, 1, filasTablaSup, colsVistaBase.length)
             .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);


  // --- 4. CONSTRUCCIÓN TABLA INFERIOR (CLONADO TÉCNICO COMPLETO) ---
  
  const filaInicioTecnica = hojaDestino.getLastRow() + 4; 
  const colInicioTecnica = 1; 
  
  const ultimaColTecnica = hojaTecnica.getLastColumn();
  if (ultimaColTecnica === 0) { return; }

  // A) Título Sección
  hojaDestino.getRange(filaInicioTecnica - 2, colInicioTecnica)
             .setValue(`DETALLE TÉCNICO COMPLETO (${nombreHojaTecnica})`)
             .setFontWeight("bold");

  // B) Copiar ANCHOS de columna (CRÍTICO: Usamos los anchos de la TÉCNICA)
  // Esto asegura que el dibujo técnico no se deforme ni se deslice.
  for (let i = 1; i <= ultimaColTecnica; i++) {
    const anchoOriginal = hojaTecnica.getColumnWidth(i);
    hojaDestino.setColumnWidth(colInicioTecnica + i - 1, anchoOriginal);
  }

  // C) Clonar CABECERA técnica
  const rangoCabeceraTec = hojaTecnica.getRange(filaCabecera, 1, 1, ultimaColTecnica);
  rangoCabeceraTec.copyTo(hojaDestino.getRange(filaInicioTecnica, colInicioTecnica));

  // D) Clonar FILAS DE DATOS
  for (let i = 0; i < indicesFilasAImprimir.length; i++) {
    const filaOriginalIdx = indicesFilasAImprimir[i];
    const filaDestinoIdx = filaInicioTecnica + 1 + i;
    
    const rangoFilaOrigen = hojaTecnica.getRange(filaOriginalIdx, 1, 1, ultimaColTecnica);
    const rangoDestino = hojaDestino.getRange(filaDestinoIdx, colInicioTecnica);
    
    rangoFilaOrigen.copyTo(rangoDestino);
    // Aseguramos que la altura de la fila técnica también se copie (importante para gráficos/planos)
    const altoFila = hojaTecnica.getRowHeight(filaOriginalIdx);
    hojaDestino.setRowHeight(filaDestinoIdx, altoFila);
  }

  // --- 5. FINALIZAR ---
  // IMPORTANTE: NO usamos autoResizeColumns aquí porque destruiría el formato de la vista técnica.
  hojaDestino.activate();
}

// ==============================================================================
// FUNCIONES AUXILIARES
// ==============================================================================

/**
 * Copia celdas específicas de una fila origen a una fila destino manteniendo TODO el formato.
 * @param {Sheet} hojaOrg Hoja de origen
 * @param {Sheet} hojaDest Hoja de destino
 * @param {number} filaOrg Número de fila en origen
 * @param {number} filaDest Número de fila en destino
 * @param {number[]} indicesCols Array de índices de columnas a copiar (ej: [1, 3, 5])
 */
function copiarCeldasDiscontinuas(hojaOrg, hojaDest, filaOrg, filaDest, indicesCols) {
  // Iteramos columna por columna para copiar formato y valor
  // Nota: Si son muchos datos (>500 filas), esto podría optimizarse, 
  // pero para hojas de impresión de pedidos es perfectamente rápido y seguro.
  for (let c = 0; c < indicesCols.length; c++) {
    const colOrigen = indicesCols[c];
    const colDestino = c + 1; // Pegamos compactado (Col 1, 2, 3...)
    
    const celdaOrigen = hojaOrg.getRange(filaOrg, colOrigen);
    const celdaDestino = hojaDest.getRange(filaDest, colDestino);
    
    celdaOrigen.copyTo(celdaDestino); // Copia todo: Valor, Color, Fuente, Bordes
  }
}