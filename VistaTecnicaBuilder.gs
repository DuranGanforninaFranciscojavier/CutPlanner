/**
 * GENERADOR DE VISTA TÉCNICA (v6 - FINAL)
 * - Pestaña color CIAN (Impresión).
 * - Clonación total de colores y datos.
 * - Sin decimales.
 */
function generarVistaTecnicaAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName("Corredera_GP-60");
  const nombreHojaTecnica = "Vista_Tecnica Corredera_GP-60";
  
  // --- CONFIGURACIÓN ---
  const INDICE_NOTA = 7; 
  const COLOR_PESTANA_IMPRESION = "#00FFFF"; // <--- CIAN ELÉCTRICO (El color de la tinta)

  const COLS_TABLA_1 = [1, 2, 3, 4, 5, 6, 15, 16, 17, 18, 19, 20, 21, 22];
  const CABECERAS_1 = ["Etiqueta", "Ancho (L)", "Alto (H)", "Hueco", "Tipo", "Alt. Testero", "Corte L", "Corte H", "M. Sup", "M. Inf", "M. Lat", "H. Lat", "H. Cen", "H. Ruedas"];
  const ANCHOS_1 = [50, 60, 60, 50, 80, 60, 60, 60, 45, 45, 45, 45, 45, 45];

  const COLS_TABLA_2 = [1, 10, 11, 12, 13, 14];
  const CABECERAS_2 = ["Etiqueta", "Cristal Ancho", "Cristal Alto", "Solape Ancho", "Solape Alto", "Atajos (MHL)"];

  // --- PREPARAR HOJA DESTINO ---
  let hojaTec = ss.getSheetByName(nombreHojaTecnica);
  if (hojaTec) ss.deleteSheet(hojaTec);
  
  // Creamos la hoja y le ponemos el COLOR INMEDIATAMENTE
  hojaTec = ss.insertSheet(nombreHojaTecnica, ss.getNumSheets());
  try {
    hojaTec.setTabColor(COLOR_PESTANA_IMPRESION); // <--- AQUI ESTÁ EL CAMBIO
  } catch(e) { /* Si falla por permisos, no rompe el script */ }

  // --- LEER DATOS (Igual que v5) ---
  const ultimaFila = hojaDatos.getLastRow();
  if (ultimaFila < 3) return;

  const rangoCabecerasOrg = hojaDatos.getRange(2, 1, 1, 22);
  const fondosCabecerasRaw = rangoCabecerasOrg.getBackgrounds()[0];
  const rangoFuente = hojaDatos.getRange(3, 1, ultimaFila - 2, 22);
  const valoresRaw = rangoFuente.getValues();
  const fondosRaw = rangoFuente.getBackgrounds();

  // --- FILTRAR ---
  const indicesValidos = [];
  for (let i = 0; i < valoresRaw.length; i++) {
    const check = valoresRaw[i][8]; 
    const id = valoresRaw[i][0];
    if (id && (check === false || check === "" || check === "FALSE")) {
      indicesValidos.push(i);
    }
  }

  if (indicesValidos.length === 0) {
    hojaTec.getRange("A1").setValue("ℹ️ Todo limpio.");
    return;
  }

  const nombreProyecto = valoresRaw[indicesValidos[0]][INDICE_NOTA - 1] || "PROYECTO SIN NOMBRE";

  // --- MAPEAR ---
  const coloresCab1 = [COLS_TABLA_1.map(idx => fondosCabecerasRaw[idx - 1])];
  const coloresCab2 = [COLS_TABLA_2.map(idx => fondosCabecerasRaw[idx - 1])];

  const limpiarDato = (valor) => (typeof valor === 'number') ? Math.round(valor) : valor;
  const valsT1 = [], fonsT1 = [];
  const valsT2 = [], fonsT2 = [];

  indicesValidos.forEach(idxFila => {
    // T1
    const filaV1 = [], filaF1 = [];
    COLS_TABLA_1.forEach(idxCol => {
      filaV1.push(limpiarDato(valoresRaw[idxFila][idxCol - 1]));
      filaF1.push(fondosRaw[idxFila][idxCol - 1]);
    });
    valsT1.push(filaV1);
    fonsT1.push(filaF1);

    // T2
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

  // Tabla 1
  const numFilas = valsT1.length;
  const numCols1 = CABECERAS_1.length;
  
  const rangoCab1 = hojaTec.getRange(3, 1, 1, numCols1);
  rangoCab1.setValues([CABECERAS_1]).setBackgrounds(coloresCab1)
           .setFontWeight("bold").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  const rangoDatos1 = hojaTec.getRange(4, 1, numFilas, numCols1);
  rangoDatos1.setValues(valsT1).setBackgrounds(fonsT1)
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  ANCHOS_1.forEach((ancho, i) => hojaTec.setColumnWidth(i + 1, ancho));

// Tabla 2
  const filaInicio2 = 4 + numFilas + 2; // <-- Ajustado: Solo 1 fila de separación entre tablas
  const numCols2 = 6;
  const CAB_2_INF = ["Etiqueta", "Ancho", "Alto", "Ancho", "Alto", "(MHL)"];

  // (Título eliminado para compactar)

// 1. Encabezado Superior (Grupos) - Fila 1
  // Usamos el color original de la columna principal de cada grupo (Indices: 9=Cristal, 11=Solape, 13=Atajo)
  hojaTec.getRange(filaInicio2, 2, 1, 2).merge().setValue("CRISTALES").setBackground(fondosCabecerasRaw[9]);
  hojaTec.getRange(filaInicio2, 4, 1, 2).merge().setValue("SOLAPE").setBackground(fondosCabecerasRaw[11]);
  hojaTec.getRange(filaInicio2, 6).setValue("ATAJOS").setBackground(fondosCabecerasRaw[13]);

  // Estilo común (Bordes y Fuente)
  hojaTec.getRange(filaInicio2, 1, 1, numCols2)
         .setFontWeight("bold").setHorizontalAlignment("center")
         .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  // Fusiones y Textos Centrados
  hojaTec.getRange(filaInicio2, 2, 1, 2).merge().setValue("CRISTALES").setHorizontalAlignment("center");
  hojaTec.getRange(filaInicio2, 4, 1, 2).merge().setValue("SOLAPE").setHorizontalAlignment("center");
  hojaTec.getRange(filaInicio2, 6).setValue("ATAJOS").setHorizontalAlignment("center");

  // 2. Encabezado Inferior (Detalles) - Fila 2 
  // AQUI aplicamos los COLORES CLONADOS (coloresCab2) para que coincidan con la hoja origen
  const rangoCab2 = hojaTec.getRange(filaInicio2 + 1, 1, 1, numCols2);
  rangoCab2.setValues([CAB_2_INF]).setBackgrounds(coloresCab2) 
           .setFontWeight("bold").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
           .setHorizontalAlignment("center");

  // 3. Datos - Fila 3 en adelante
  const rangoDatos2 = hojaTec.getRange(filaInicio2 + 2, 1, numFilas, numCols2);
  rangoDatos2.setValues(valsT2).setBackgrounds(fonsT2)
             .setHorizontalAlignment("center").setVerticalAlignment("middle")
             .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  console.log("✅ Vista Técnica (Pestaña CIAN): " + nombreProyecto);
}