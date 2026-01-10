/**
 * GENERADOR DE VISTA TÉCNICA (VERSIÓN 2.0 - FORMATO "CLAVEL ROJO")
 * - Incluye columna "Hueco".
 * - Aplica bordes, centrados y anchos de columna específicos.
 */
function generarVistaTecnicaAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName("Corredera_GP-60");
  const nombreHojaTecnica = "Vista_Tecnica Corredera_GP-60";
  
  // --- 1. CONFIGURACIÓN DE COLUMNAS ---
  // AÑADIDO: Columna 4 (Hueco)
  const COLS_TABLA_1 = [1, 2, 3, 4, 5, 6, 15, 16, 17, 18, 19, 20, 21, 22];
  const CABECERAS_1 = ["Etiqueta", "Ancho (L)", "Alto (H)", "Hueco", "Tipo", "Alt. Testero", "Corte L", "Corte H", "M. Sup", "M. Inf", "M. Lat", "H. Lat", "H. Cen", "H. Ruedas"];
  // Anchos aproximados para Tabla 1 (Suman ~900px, cabe en A4 Horizontal)
  const ANCHOS_1 = [50, 60, 60, 50, 80, 60, 65, 65, 65, 65, 65, 65, 65, 65];

  const COLS_TABLA_2 = [1, 10, 11, 12, 13, 14];
  const CABECERAS_2 = ["Etiqueta", "Cristal Ancho", "Cristal Alto", "Solape Ancho", "Solape Alto", "Atajos (MHL)"];
  const ANCHOS_2 = [50, 90, 90, 90, 90, 90];

  // --- 2. PREPARAR HOJA ---
  let hojaTec = ss.getSheetByName(nombreHojaTecnica);
  if (hojaTec) ss.deleteSheet(hojaTec);
  hojaTec = ss.insertSheet(nombreHojaTecnica, ss.getNumSheets());

  // --- 3. OBTENER DATOS PENDIENTES ---
  const ultimaFila = hojaDatos.getLastRow();
  if (ultimaFila < 3) return; 
  
  const datosRaw = hojaDatos.getRange(3, 1, ultimaFila - 2, 22).getValues();
  // Filtramos por Check (Columna I, índice 8) == FALSE
  const filasFiltradas = datosRaw.filter(fila => {
    const check = fila[8]; 
    const id = fila[0];
    return id && (check === false || check === "" || check === "FALSE");
  });

  if (filasFiltradas.length === 0) {
    hojaTec.getRange("A1").setValue("ℹ️ Todo limpio. No hay trabajos pendientes.");
    return;
  }

  // --- 4. MAPEAR DATOS ---
  const datosTabla1 = filasFiltradas.map(fila => COLS_TABLA_1.map(idx => fila[idx - 1]));
  const datosTabla2 = filasFiltradas.map(fila => COLS_TABLA_2.map(idx => fila[idx - 1]));

  // --- 5. PINTAR TABLA SUPERIOR (DATOS) ---
  const numFilas = datosTabla1.length;
  const numCols1 = CABECERAS_1.length;

  // Título Principal
  hojaTec.getRange("A1").setValue("VISTA TÉCNICA - CORREDERA 2 HOJAS")
         .setFontWeight("bold").setFontSize(14);
  
  // Cabeceras y Datos
  const rangoTabla1 = hojaTec.getRange(3, 1, numFilas + 1, numCols1);
  // Escribimos Cabeceras (Fila 3) + Datos (Fila 4 en adelante)
  hojaTec.getRange(3, 1, 1, numCols1).setValues([CABECERAS_1]);
  hojaTec.getRange(4, 1, numFilas, numCols1).setValues(datosTabla1);
  
  // ESTILO TABLA 1
  rangoTabla1.setHorizontalAlignment("center").setVerticalAlignment("middle");
  rangoTabla1.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  hojaTec.getRange(3, 1, 1, numCols1).setBackground("#E0E0E0").setFontWeight("bold"); // Gris cabecera

  // APLICAR ANCHOS TABLA 1
  ANCHOS_1.forEach((ancho, i) => hojaTec.setColumnWidth(i + 1, ancho));

  // --- 6. PINTAR TABLA INFERIOR (DETALLES) ---
  const filaInicio2 = 4 + numFilas + 3; // Dejamos espacio
  const numCols2 = CABECERAS_2.length;
  
  // Título Sección 2
  hojaTec.getRange(filaInicio2 - 2, 1).setValue("DETALLES (Cristales y Solapes)").setFontWeight("bold");

  const rangoTabla2 = hojaTec.getRange(filaInicio2, 1, numFilas + 1, numCols2);
  hojaTec.getRange(filaInicio2, 1, 1, numCols2).setValues([CABECERAS_2]);
  hojaTec.getRange(filaInicio2 + 1, 1, numFilas, numCols2).setValues(datosTabla2);

  // ESTILO TABLA 2
  rangoTabla2.setHorizontalAlignment("center").setVerticalAlignment("middle");
  rangoTabla2.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  hojaTec.getRange(filaInicio2, 1, 1, numCols2).setBackground("#E0E0E0").setFontWeight("bold");
  
  // (Opcional) Re-aplicar anchos para Tabla 2 si fueran muy distintos, 
  // pero como comparten columnas, mandan los anchos de arriba. 
  // Solo ajustamos las columnas A-F que son las que usa la Tabla 2.

  console.log("✅ Vista Técnica Refinada generada.");
}