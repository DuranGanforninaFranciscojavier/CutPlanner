// ==============================================================================
// SECCIÓN 1: CLONADOR DE TABLA PARA IMPRESIÓN (Plan_Cortes -> cortes_impresión)
// ==============================================================================

/**
 * CLONADOR DE TABLA CON SELECCIÓN EXPLÍCITA DE COLUMNAS
 * ----------------------------------------------------
 * REFLEXIÓN: Detectar columnas ocultas automáticamente en Apps Script es poco fiable.
 * SOLUCIÓN: Definimos manualmente qué columnas queremos en la impresión.
 * * Pasos:
 * 1. Define la lista exacta de columnas a imprimir (indicesVisibles).
 * 2. Busca la hoja "Plan_Cortes" y obtiene los anchos de esas columnas específicas.
 * 3. Crea/Resetea la hoja "cortes_impresión".
 * 4. Extrae los datos y formatos SOLO de la lista definida.
 * 5. Divide y pega en dos mitades.
 */
function dividirTablaParaImpresionClonada() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaOrigen = "Plan_Cortes";
  const nombreHojaDestino = "cortes_impresión";
  const filaCabecera = 1;
  const primeraFilaDatos = 2;

  // --- 1. SETUP Y VALIDACIONES ---
  const hojaOrigen = ss.getSheetByName(nombreHojaOrigen);
  if (!hojaOrigen) {
    SpreadsheetApp.getUi().alert(`Error: No se encuentra la hoja "${nombreHojaOrigen}".`);
    return;
  }

  const ultimaFila = hojaOrigen.getLastRow();
  if (ultimaFila < primeraFilaDatos) {
     SpreadsheetApp.getUi().alert(`La hoja "${nombreHojaOrigen}" no tiene datos suficientes.`);
     return;
  }

  // --- 2. DEFINICIÓN EXPLÍCITA DE COLUMNAS A IMPRIMIR ---
  // ESTA ES LA CLAVE: No intentamos detectar, definimos lo que queremos.
  // Columnas típicas de taller: Serie(2), Perfil(3), Tipo(5), Cortes(7), Resto(8)
  // ADAPTA ESTA LISTA SI QUIERES OTRAS.
  const indicesVisibles = [2, 3, 7, 8]; 
  
  const numColsVisibles = indicesVisibles.length;
  const anchosVisibles = [];

  // Obtenemos los anchos originales SOLO de las columnas seleccionadas
  for (let i = 0; i < numColsVisibles; i++) {
    const indiceReal = indicesVisibles[i];
    // Validamos que la columna exista
    if (indiceReal > hojaOrigen.getLastColumn()) {
       SpreadsheetApp.getUi().alert(`Error: La columna ${indiceReal} no existe en el origen.`);
       return;
    }
    anchosVisibles.push(hojaOrigen.getColumnWidth(indiceReal));
  }


  // --- 3. PREPARAR HOJA DESTINO ---
  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (hojaDestino) ss.deleteSheet(hojaDestino);
  hojaDestino = ss.insertSheet(nombreHojaDestino, ss.getNumSheets());


  // --- 4. CÁLCULOS DE DIVISIÓN ---
  const totalFilasDatos = ultimaFila - primeraFilaDatos + 1;
  const filasPorMitad = Math.ceil(totalFilasDatos / 2);
  const filaFinIzq = primeraFilaDatos + filasPorMitad - 1;
  const filaInicioDer = filaFinIzq + 1;


  // --- 5. EXTRACCIÓN Y PEGADO DE DATOS FILTRADOS ---
  // (Utiliza las funciones auxiliares definidas abajo)

  // A) CABECERA
  const datosCabecera = extraerDatosFiltrados(hojaOrigen, 1, 1, indicesVisibles);
  pegarDatosFiltrados(hojaDestino, 1, 1, datosCabecera);
  
  // B) MITAD IZQUIERDA
  const datosIzq = extraerDatosFiltrados(hojaOrigen, primeraFilaDatos, filasPorMitad, indicesVisibles);
  pegarDatosFiltrados(hojaDestino, 2, 1, datosIzq);

  // C) MITAD DERECHA
  const colInicioDer = numColsVisibles + 2; // Dejamos 1 columna de separación
  
  pegarDatosFiltrados(hojaDestino, 1, colInicioDer, datosCabecera);
  
  const filasRestantes = totalFilasDatos - filasPorMitad;
  if (filasRestantes > 0) {
    const datosDer = extraerDatosFiltrados(hojaOrigen, filaInicioDer, filasRestantes, indicesVisibles);
    pegarDatosFiltrados(hojaDestino, 2, colInicioDer, datosDer);
  }


  // --- 6. APLICAR ANCHOS DE COLUMNA ---
  for (let i = 0; i < numColsVisibles; i++) {
    const ancho = anchosVisibles[i];
    const indiceDestinoBase1 = i + 1;
    // Aplicar a lado izquierdo
    hojaDestino.setColumnWidth(indiceDestinoBase1, ancho);
    // Aplicar a lado derecho
    hojaDestino.setColumnWidth(colInicioDer + i, ancho);
  }
  // Ajustar separador
  hojaDestino.setColumnWidth(numColsVisibles + 1, 20);


  // --- 7. ESTILO FINAL DE CABECERAS ---
  // Seleccionamos toda la fila 1 de la hoja de destino
  const rangoCabeceras = hojaDestino.getRange(1, 1, 1, hojaDestino.getLastColumn());
  // Aplicamos fondo gris y fuente blanca y negrita para que resalte
  rangoCabeceras.setBackground("#434343") // Gris oscuro
                .setFontColor("#ffffff") // Color blanco
                .setFontWeight("bold");

  hojaDestino.activate();
}


/**
 * FUNCIÓN AUXILIAR (NECESARIA): Extrae datos y formatos solo de las columnas indicadas.
 */
function extraerDatosFiltrados(hoja, filaInicio, numFilas, indicesVisibles) {
  const ultimaColOrigen = hoja.getLastColumn();
  // Leemos TODO el rango en memoria primero
  const rangoCompleto = hoja.getRange(filaInicio, 1, numFilas, ultimaColOrigen);
  
  const valoresRaw = rangoCompleto.getDisplayValues(); 
  const fondosRaw = rangoCompleto.getBackgrounds();
  const fuentesRaw = rangoCompleto.getFontWeights();
  const alineacionRaw = rangoCompleto.getHorizontalAlignments();

  // Matrices de destino
  const valoresFiltrados = [];
  const fondosFiltrados = [];
  const fuentesFiltradas = [];
  const alineacionFiltrada = [];

  for (let f = 0; f < numFilas; f++) {
    const filaVal = [], filaFon = [], filaFue = [], filaAli = [];
    // Iteramos solo sobre los índices que HEMOS DEFINIDO MANUALMENTE
    for (let i = 0; i < indicesVisibles.length; i++) {
      const colIdxArray = indicesVisibles[i] - 1; // Convertir índice base-1 a base-0
      
      // Extraemos el dato exacto de esa columna
      filaVal.push(valoresRaw[f][colIdxArray]);
      filaFon.push(fondosRaw[f][colIdxArray]);
      filaFue.push(fuentesRaw[f][colIdxArray]);
      filaAli.push(alineacionRaw[f][colIdxArray]);
    }
    valoresFiltrados.push(filaVal);
    fondosFiltrados.push(filaFon);
    fuentesFiltradas.push(filaFue);
    alineacionFiltrada.push(filaAli);
  }

  return {
    numFilas: numFilas,
    numCols: indicesVisibles.length,
    valores: valoresFiltrados,
    fondos: fondosFiltrados,
    fuentes: fuentesFiltradas,
    alineacion: alineacionFiltrada
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