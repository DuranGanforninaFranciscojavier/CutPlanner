// ==============================================================================
// SECCIÓN 1: CLONADOR DE TABLA PARA IMPRESIÓN (Plan_Cortes -> cortes_impresión)
// ==============================================================================
/**
 * CLONADOR HÍBRIDO: FILTRA, IMPRIME Y MARCA COMO HECHO
 * 1. Filtra filas pendientes (Hecha == FALSE).
 * 2. Genera hoja con TUS anchos [43, 120, 100, 50].
 * 3. Marca los originales como TRUE (Hecho).
 */
function dividirTablaParaImpresionClonada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaOrigen = "Plan_Cortes";
  const nombreHojaDestino = "cortes_impresión";
  const primeraFilaDatos = 2;

  // --- 1. LEER ORIGEN Y FILTRAR ---
  const hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  if (!hojaOrigen) return;

  const ultimaFila = hojaOrigen.getLastRow();
  if (ultimaFila < primeraFilaDatos) {
    SpreadsheetApp.getUi().alert("⚠️ Plan_Cortes vacío.");
    return;
  }

  // Detectar filas PENDIENTES (Columna I = 9)
  const rangoCheck = hojaOrigen.getRange(primeraFilaDatos, 9, ultimaFila - primeraFilaDatos + 1, 1).getValues();
  const indicesFilasValidas = []; // Guardará el número de fila REAL (Base 1)
  
  for (let i = 0; i < rangoCheck.length; i++) {
    const hecha = rangoCheck[i][0];
    // Si NO está hecha (FALSE o vacío), la procesamos
    if (hecha !== true && String(hecha).toUpperCase() !== "TRUE") {
      indicesFilasValidas.push(primeraFilaDatos + i);
    }
  }

  // --- 2. PREPARAR DESTINO ---
  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (hojaDestino) ss.deleteSheet(hojaDestino);
  hojaDestino = ss.insertSheet(nombreHojaDestino, ss.getNumSheets());
  try { hojaDestino.setTabColor("#00FFFF"); } catch(e) {}

  if (indicesFilasValidas.length === 0) {
    hojaDestino.getRange("A1").setValue("✅ Todo cortado. Historial limpio.");
    return;
  }

  // --- 3. CONFIGURACIÓN VISUAL (TUYA) ---
  const indicesVisibles = [2, 3, 7, 8]; // Serie, Perfil, Cortes, Resto
  const anchosCompactos = [44, 120, 100, 50]; // TUS ANCHOS
  const numColsVisibles = indicesVisibles.length;

  // --- 4. EXTRACCIÓN Y PEGADO ---
  const totalItems = indicesFilasValidas.length;
  const mitad = Math.ceil(totalItems / 2);
  
  const filasIzq = indicesFilasValidas.slice(0, mitad);
  const filasDer = indicesFilasValidas.slice(mitad);

  // A) Cabecera (Fila 1 siempre)
  const datosCabecera = extraerDatosFiltradosPorIndices(hojaOrigen, [1], indicesVisibles);
  pegarDatosFiltrados(hojaDestino, 1, 1, datosCabecera);

  // B) Mitad Izquierda
  if (filasIzq.length > 0) {
    const datosIzq = extraerDatosFiltradosPorIndices(hojaOrigen, filasIzq, indicesVisibles);
    pegarDatosFiltrados(hojaDestino, 2, 1, datosIzq);
  }

  // C) Mitad Derecha
  const colInicioDer = numColsVisibles + 2;
  pegarDatosFiltrados(hojaDestino, 1, colInicioDer, datosCabecera);

  if (filasDer.length > 0) {
    const datosDer = extraerDatosFiltradosPorIndices(hojaOrigen, filasDer, indicesVisibles);
    pegarDatosFiltrados(hojaDestino, 2, colInicioDer, datosDer);
  }

  // --- 5. FORMATO FINAL ---
  for (let i = 0; i < numColsVisibles; i++) {
    const ancho = anchosCompactos[i];
    hojaDestino.setColumnWidth(i + 1, ancho);              // Izquierda
    hojaDestino.setColumnWidth(colInicioDer + i, ancho);   // Derecha
  }
  hojaDestino.setColumnWidth(numColsVisibles + 1, 20); // Separador

  const rangoCabeceras = hojaDestino.getRange("1:1");
  rangoCabeceras.setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
  
  // --- CAMBIO: AJUSTE DE TEXTO Y TAMAÑO DE FUENTE 11 ---
  hojaDestino.getDataRange()
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setFontSize(11); // <--- TAMAÑO 11 IMPUESTO

  // --- 6. MARCAR COMO HECHO EN ORIGEN (CRÍTICO) ---
  // Recorremos la lista de lo que acabamos de imprimir y ponemos check TRUE
  indicesFilasValidas.forEach(fila => {
    hojaOrigen.getRange(fila, 9).setValue(true); // Col 9 = Hecha
  });
  
  hojaDestino.activate();
}
/**
 * NUEVA AUXILIAR: Extrae filas sueltas y REDONDEA números.
 */
function extraerDatosFiltradosPorIndices(hoja, indicesFilas, indicesCols) {
  const valoresFiltrados = [], fondosFiltrados = [];
  const fuentesFiltradas = [], alineacionFiltrada = [];

  indicesFilas.forEach(filaNum => {
    const filaVal = [], filaFon = [], filaFue = [], filaAli = [];
    indicesCols.forEach(colNum => {
      const celda = hoja.getRange(filaNum, colNum);
      
      // --- CORRECCIÓN REDONDEO ---
      let valor = celda.getValue(); // Usamos valor crudo, no DisplayValue
      if (typeof valor === 'number') {
        valor = Math.round(valor); // Hachazo a los decimales
      }
      
      filaVal.push(valor);
      filaFon.push(celda.getBackground());
      filaFue.push(celda.getFontWeight());
      filaAli.push(celda.getHorizontalAlignment());
    });
    valoresFiltrados.push(filaVal); fondosFiltrados.push(filaFon);
    fuentesFiltradas.push(filaFue); alineacionFiltrada.push(filaAli);
  });

  return {
    numFilas: indicesFilas.length, numCols: indicesCols.length,
    valores: valoresFiltrados, fondos: fondosFiltrados,
    fuentes: fuentesFiltradas, alineacion: alineacionFiltrada
  };
}

/**
 * FUNCIÓN AUXILIAR (NECESARIA): Pega los datos extraídos en el destino.
 */
function pegarDatosFiltrados(hojaDest, filaIni, colIni, paqueteDatos) {
  if (paqueteDatos.numFilas === 0) return;
  const rangoDestino = hojaDest.getRange(filaIni, colIni, paqueteDatos.numFilas, paqueteDatos.numCols);
  rangoDestino.setValues(paqueteDatos.valores);
  rangoDestino.setBackgrounds(paqueteDatos.fondos);
  rangoDestino.setFontWeights(paqueteDatos.fuentes);
  rangoDestino.setHorizontalAlignments(paqueteDatos.alineacion);
}
/**
 * GENERADOR DE IMPRESIÓN COMBINADA (HÍBRIDO: MANUAL + FUENTE GIGANTE)
 * - Fuente: 18pt (Gigante) para legibilidad máxima.
 * - Columnas: DEFINIDAS MANUALMENTE POR CÓDIGO (Tú controlas el ancho).
 * - Ajuste: Mantiene el recorte de sobrantes para que el PDF haga Zoom al máximo.
 */
function generarImpresionCombinada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaDestino = "Impresion_Combinada";

  // 1. LOCALIZAR
  const hojas = ss.getSheets();
  const hojaTecnica = hojas.find(h => h.getName().startsWith("Vista_Tecnica"));
  const hojaCortes = ss.getSheetByName("cortes_impresión");

  if (!hojaTecnica || !hojaCortes) {
    SpreadsheetApp.getUi().alert("⚠️ Faltan hojas. Ejecuta '▶️ PRODUCCIÓN' primero.");
    return;
  }

  // 2. CLONAR
  let hojaDest = ss.getSheetByName(nombreHojaDestino);
  if (hojaDest) ss.deleteSheet(hojaDest);
  hojaDest = hojaTecnica.copyTo(ss);
  hojaDest.setName(nombreHojaDestino);
  try { hojaDest.setTabColor("#FF00FF"); } catch(e) {}

  // 3. PEGAR CORTES
  const lastRowTec = hojaDest.getLastRow();
  const filaInicioCortes = lastRowTec + 2; // 1 fila de aire

  // Título Separador
  hojaDest.getRange(filaInicioCortes, 1, 1, 8).merge()
          .setValue("PLAN DE CORTES Y TALLER")
          .setBackground("#434343").setFontColor("white")
          .setFontWeight("bold").setHorizontalAlignment("center")
          .setFontSize(20);

  const datosCortes = hojaCortes.getDataRange();
  datosCortes.copyTo(hojaDest.getRange(filaInicioCortes + 1, 1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  datosCortes.copyTo(hojaDest.getRange(filaInicioCortes + 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // --- ESTRATEGIA VISUAL ---
  const rangoTotal = hojaDest.getDataRange();
  
  // A. FUENTES GIGANTES
  rangoTotal.setFontSize(18);              // Base gigante
  hojaDest.getRange("A1").setFontSize(26); // Título Proyecto

  // =======================================================
  //      👇👇👇 ZONA DE CONTROL DE ANCHOS (PÍXELES) 👇👇👇
  // =======================================================
  // Modifica estos números para ajustar tu diseño. 
  // Al usar fuente 18, los valores deben ser un poco más altos de lo normal.
  
  hojaDest.setColumnWidth(1, 109);   // A: Etiqueta
  hojaDest.setColumnWidth(2, 117);  // B: Ancho / Perfil (Necesita espacio)
  hojaDest.setColumnWidth(3, 170);   // C: Alto / Medida
  hojaDest.setColumnWidth(4, 81);   // D: Hueco
  hojaDest.setColumnWidth(5, 124);   // E: Tipo
  hojaDest.setColumnWidth(6, 134);   // F: Tester / Extra
  hojaDest.setColumnWidth(7, 117);   // G: Tester / Extra
  hojaDest.setColumnWidth(8, 170);   // H: Tester / Extra
  hojaDest.setColumnWidth(9, 84);   // i: Tester / Extra
  hojaDest.setColumnWidth(13, 76);   // n: Tester / Extra

  // Columnas de Medidas de Corte (G en adelante)
  // Ajusta según necesites más o menos espacio para los números
  for (let i = 10; i <= 20; i++) {
     try { hojaDest.setColumnWidth(i, 75); } catch(e) {}
  }
  // =======================================================

  // B. AJUSTAR FILAS (Automático para que el texto 18 quepa verticalmente)
  hojaDest.autoResizeRows(1, hojaDest.getLastRow());

  // 4. TRIM (Borrar vacíos - CRÍTICO PARA EL ZOOM)
  const ultFilaDatos = hojaDest.getLastRow();
  const ultColDatos = hojaDest.getLastColumn();
  
  if (hojaDest.getMaxRows() > ultFilaDatos) {
    hojaDest.deleteRows(ultFilaDatos + 1, hojaDest.getMaxRows() - ultFilaDatos);
  }
  if (hojaDest.getMaxColumns() > ultColDatos) {
    hojaDest.deleteColumns(ultColDatos + 1, hojaDest.getMaxColumns() - ultColDatos);
  }

  // 5. ALINEACIÓN FINAL
  rangoTotal.setVerticalAlignment("middle").setHorizontalAlignment("center");
  // Wrap activado por si alguna columna manual es muy estrecha y queremos que salte línea
  rangoTotal.setWrap(true); 
  
  hojaDest.setHiddenGridlines(true);

  return hojaDest.getName();
}