/**
 * GENERADOR DE VISTA TÉCNICA (v7 - FINAL LIMPIA)
 * - Pestaña CIAN.
 * - Clonación exacta de colores y encabezados.
 * - Doble tabla con encabezados agrupados.
 */
function generarVistaTecnicaAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName("Corredera_GP-60");
  const nombreHojaTecnica = "Vista_Tecnica Corredera_GP-60";
  
  // --- CONFIGURACIÓN ---
  const INDICE_NOTA = 7; 
  const COLOR_PESTANA_IMPRESION = "#00FFFF"; 

  const COLS_TABLA_1 = [1, 2, 3, 4, 5, 6, 15, 16, 17, 18, 19, 20, 21, 22];
  const CABECERAS_1 = ["Etiqueta", "Ancho(L)", "Alto(H)", "Hueco", "Tipo", "Alt.Testero", "Corte L", "Corte H", "M.Sup", "M.Inf", "M.Lat", "H.Lat", "H.Cen", "H.Rue"];
  // Anchos ajustados para compactar
  const ANCHOS_1 = [50, 60, 60, 50, 80, 60, 60, 60, 47, 45, 45, 45, 45, 45];

  const COLS_TABLA_2 = [1, 10, 11, 12, 13, 14];
  const CABECERAS_2 = ["Etiqueta", "Cristal Ancho", "Cristal Alto", "Solape Ancho", "Solape Alto", "Atajos (MHL)"];

  // --- PREPARAR HOJA DESTINO ---
  let hojaTec = ss.getSheetByName(nombreHojaTecnica);
  if (hojaTec) ss.deleteSheet(hojaTec);
  hojaTec = ss.insertSheet(nombreHojaTecnica, ss.getNumSheets());
  try { hojaTec.setTabColor(COLOR_PESTANA_IMPRESION); } catch(e) {}

 // --- LEER DATOS ---
  const ultimaFila = hojaDatos.getLastRow();
  if (ultimaFila < 3) return;

  // LEEMOS HASTA LA COLUMNA 23 (W) para incluir el Nombre del Proyecto
  // Nota: Si la hoja no tiene col 23, getRange podría fallar, pero asumimos que ya la has creado.
  const numColsLectura = Math.max(23, hojaDatos.getLastColumn());
  
  const rangoCabecerasOrg = hojaDatos.getRange(2, 1, 1, 22); // Encabezados visuales (solo hasta V)
  const fondosCabecerasRaw = rangoCabecerasOrg.getBackgrounds()[0];
  
  const rangoFuente = hojaDatos.getRange(3, 1, ultimaFila - 2, numColsLectura);
  const valoresRaw = rangoFuente.getValues();
  const fondosRaw = rangoFuente.getBackgrounds();

  // --- FILTRAR ---
  const indicesValidos = [];
  for (let i = 0; i < valoresRaw.length; i++) {
    const check = valoresRaw[i][8]; // Col I
    const id = valoresRaw[i][0];    // Col A
    if (id && (check === false || check === "" || check === "FALSE")) {
      indicesValidos.push(i);
    }
  }

  if (indicesValidos.length === 0) {
    hojaTec.getRange("A1").setValue("ℹ️ Todo limpio.");
    return;
  }

  // OBTENER NOMBRE DEL PROYECTO (Columna 23 -> Índice 22)
  // Leemos de la primera fila válida encontrada
  let nombreProyecto = valoresRaw[indicesValidos[0]][22]; 
  if (!nombreProyecto) nombreProyecto = "PROYECTO (Col W Vacía)";

  // --- MAPEAR ---
  const coloresCab1 = [COLS_TABLA_1.map(idx => fondosCabecerasRaw[idx - 1])];
  const coloresCab2 = [COLS_TABLA_2.map(idx => fondosCabecerasRaw[idx - 1])];

  const limpiarDato = (valor) => (typeof valor === 'number') ? Math.round(valor) : valor;
  const valsT1 = [], fonsT1 = [];
  const valsT2 = [], fonsT2 = [];

  indicesValidos.forEach(idxFila => {
    const filaV1 = [], filaF1 = [];
    COLS_TABLA_1.forEach(idxCol => {
      filaV1.push(limpiarDato(valoresRaw[idxFila][idxCol - 1]));
      filaF1.push(fondosRaw[idxFila][idxCol - 1]);
    });
    valsT1.push(filaV1);
    fonsT1.push(filaF1);

    const filaV2 = [], filaF2 = [];
    COLS_TABLA_2.forEach(idxCol => {
      filaV2.push(limpiarDato(valoresRaw[idxFila][idxCol - 1]));
      filaF2.push(fondosRaw[idxFila][idxCol - 1]);
    });
    valsT2.push(filaV2);
    fonsT2.push(filaF2);
  });

  // --- PINTAR ---
  hojaTec.getRange("A1").setValue(nombreProyecto.toUpperCase()).setFontWeight("bold").setFontSize(16);
// 1. TÍTULO PROYECTO (Centrado en A1:N1)
  hojaTec.getRange("A1:N1").merge()
         .setValue(nombreProyecto.toUpperCase())
         .setHorizontalAlignment("center").setFontWeight("bold").setFontSize(16);

  // === TABLA 1 (SUPERIOR) ===
  const numFilas = valsT1.length;
  const numCols1 = CABECERAS_1.length;

  // 2. ENCABEZADOS DE GRUPO (FILA 2)
  // Medidas Corte (G-H), Marco (I-K), Hoja (L-N)
  hojaTec.getRange(2, 7, 1, 2).merge().setValue("Medidas Corte").setBackground("#76A5AF"); // Azulado
  hojaTec.getRange(2, 9, 1, 3).merge().setValue("Marco").setBackground("#999999");         // Gris
  hojaTec.getRange(2, 12, 1, 3).merge().setValue("Hoja").setBackground("#999999");         // Gris Oscuro
  
  // Formato Fila 2 (Bordes y Fuente)
  hojaTec.getRange(2, 7, 1, 8)
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // 3. ENCABEZADOS ESPECÍFICOS (FILA 3)
  const rangoCab1 = hojaTec.getRange(3, 1, 1, numCols1);
  rangoCab1.setValues([CABECERAS_1]).setBackgrounds(coloresCab1)
           .setFontWeight("bold").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
           .setHorizontalAlignment("center");
  
  // 4. DATOS (FILA 4 EN ADELANTE)
  const rangoDatos1 = hojaTec.getRange(4, 1, numFilas, numCols1);
  rangoDatos1.setValues(valsT1).setBackgrounds(fonsT1)
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // --- NUEVO: COLUMNAS EN NEGRITA (Etiqueta, Hueco, Tipo) ---
  // Indices (Base 1): Etiqueta=1, Hueco=4, Tipo=5
  hojaTec.getRange(4, 1, numFilas, 1).setFontWeight("bold");
  hojaTec.getRange(4, 4, numFilas, 1).setFontWeight("bold");
  hojaTec.getRange(4, 5, numFilas, 1).setFontWeight("bold");

  ANCHOS_1.forEach((ancho, i) => hojaTec.setColumnWidth(i + 1, ancho));

  // === TABLA 2 (Compacta y Doble Encabezado) ===
  const filaInicio2 = 4 + numFilas + 2;
  const numCols2 = 6;
  const CAB_2_INF = ["Etiqueta", "Ancho", "Alto", "Ancho", "Alto", "(MHL)"];

  // 1. Encabezado Superior (Grupos)
  // Usamos el color original de la columna principal de cada grupo
  hojaTec.getRange(filaInicio2, 2, 1, 2).merge().setValue("CRISTALES").setBackground("#8e7cc3");
  hojaTec.getRange(filaInicio2, 4, 1, 2).merge().setValue("SOLAPE").setBackground("#e06666");
  hojaTec.getRange(filaInicio2, 6).setValue("ATAJOS").setBackground(fondosCabecerasRaw[13]);

  // Estilo común
  hojaTec.getRange(filaInicio2, 1, 1, numCols2)
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // 2. Encabezado Inferior (Detalles)
  const rangoCab2 = hojaTec.getRange(filaInicio2 + 1, 1, 1, numCols2);
  rangoCab2.setValues([CAB_2_INF]).setBackgrounds(coloresCab2) 
           .setFontWeight("bold").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
           .setHorizontalAlignment("center");

  // 3. Datos
  const rangoDatos2 = hojaTec.getRange(filaInicio2 + 2, 1, numFilas, numCols2);
  rangoDatos2.setValues(valsT2).setBackgrounds(fonsT2)
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
// --- AJUSTE FINAL DE FUENTE ---
  // 1. Ponemos TODO el contenido a tamaño 11
  hojaTec.getDataRange().setFontSize(11);
  
  // 2. Restauramos el Título Principal (A1) a tamaño 16 para que destaque
  hojaTec.getRange("A1").setFontSize(16);

  console.log("✅ Vista Técnica (Pestaña CIAN): " + nombreProyecto);
}