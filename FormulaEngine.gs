/**
 * ==============================================================================
 * ARCHIVO: FormulaEngine.gs
 * MOTOR DE CÁLCULO Y DETECCIÓN DE CAMBIOS
 * Mapeo Dinámico por Encabezados (Fila 1) y Fórmulas de Carpintería
 * ==============================================================================
 */

/**
 * Disparador automático que detecta cambios en las celdas de entrada.
 * Escucha las columnas de cotas y tipos en la hoja de trabajo y recalcula la fila.
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Edición en la hoja de datos (a partir de FILA_INICIO_DATOS)
  if (sheetName === HOJA_DATOS_NOMBRE && row >= FILA_INICIO_DATOS) {
    const colMap = getHeaderMap(sheet, FILA_CABECERA);

    // Identificar columnas que disparan el recálculo
    const colsGatillo = [
      colMap.getCol("Ancho_Luz"),
      colMap.getCol("Alto_Luz"),
      colMap.getCol("Ancho_Corte"),
      colMap.getCol("Alto_Corte"),
      colMap.getCol("Tipo_Perfil"),
      colMap.getCol("Altura_Testero")
    ].filter(c => c !== null);

    if (colsGatillo.includes(col)) {
      const config = leerConfiguracion(sheet.getParent());
      calcularFila(sheet, row, colMap, config);
    }
  }
}

/**
 * Recorre toda la hoja y fuerza el recálculo de todas las filas con datos.
 */
function actualizarTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOJA_DATOS_NOMBRE);
  if (!sheet) {
    console.warn(`No se encontró la hoja "${HOJA_DATOS_NOMBRE}"`);
    return;
  }

  const ultimaFila = sheet.getLastRow();
  if (ultimaFila < FILA_INICIO_DATOS) {
    SpreadsheetApp.getActive().toast("No se encontraron filas con datos para actualizar.");
    return;
  }

  const colMap = getHeaderMap(sheet, FILA_CABECERA);
  const config = leerConfiguracion(ss);
  let filasActualizadas = 0;

  for (let i = FILA_INICIO_DATOS; i <= ultimaFila; i++) {
    const exito = calcularFila(sheet, i, colMap, config);
    if (exito) filasActualizadas++;
  }

  SpreadsheetApp.getActive().toast(`Se han recalculado ${filasActualizadas} ventanas.`, "🔄 Recálculo Completo");
}

/**
 * Lógica principal de cálculo de una fila.
 * Maneja Luz vs Corte, descuentos de Configuración, perfiles y cristales.
 */
function calcularFila(sheet, fila, colMap, config) {
  if (!colMap) colMap = getHeaderMap(sheet, FILA_CABECERA);
  if (!config) config = leerConfiguracion(sheet.getParent());

  const lastCol = sheet.getLastColumn();
  const rowValues = sheet.getRange(fila, 1, 1, lastCol).getValues()[0];

  function getVal(colName) {
    const colIdx = colMap.getCol(colName);
    if (!colIdx || colIdx > rowValues.length) return "";
    return rowValues[colIdx - 1];
  }

  function setVal(colName, value) {
    const colIdx = colMap.getCol(colName);
    if (colIdx) {
      sheet.getRange(fila, colIdx).setValue(value);
    }
  }

  // 1. LECTURA DE VALORES BRUTOS
  const anchoLuzRaw = getVal("Ancho_Luz");
  const altoLuzRaw = getVal("Alto_Luz");
  const anchoCorteRaw = getVal("Ancho_Corte");
  const altoCorteRaw = getVal("Alto_Corte");
  const tipoPerfil = String(getVal("Tipo_Perfil") || "").trim();
  const testeroRaw = getVal("Altura_Testero");

  const anchoLuz = parseFloat(String(anchoLuzRaw).replace(",", "."));
  const altoLuz = parseFloat(String(altoLuzRaw).replace(",", "."));
  let anchoCorte = parseFloat(String(anchoCorteRaw).replace(",", "."));
  let altoCorte = parseFloat(String(altoCorteRaw).replace(",", "."));
  const T = parseFloat(String(testeroRaw).replace(",", "."));

  // Descuentos desde Configuración
  const descVentana = Number(config.Descuento_Ventana || config.Descuento_Ventana || 0);
  const descCristal = Number(config.Descuento_Cristal || 0);

  // 2. LÓGICA LUZ VS CORTE
  // Si se introduce Ancho_Luz válido, calculamos Ancho_Corte
  if (!isNaN(anchoLuz) && anchoLuz > 0) {
    anchoCorte = anchoLuz - descVentana;
    setVal("Ancho_Corte", anchoCorte);
  }
  // Si se introduce Alto_Luz válido, calculamos Alto_Corte
  if (!isNaN(altoLuz) && altoLuz > 0) {
    altoCorte = altoLuz - descVentana;
    setVal("Alto_Corte", altoCorte);
  }

  // Medidas efectivas de corte para la carpintería
  const L = anchoCorte;
  const H = altoCorte;

  // Validación: Se requieren L, H y Tipo_Perfil
  if (isNaN(L) || L <= 0 || isNaN(H) || H <= 0 || !tipoPerfil) {
    return false;
  }

  // 3. CÁLCULO DE PERFILES (MARCOS Y HOJAS)
  let marcoSup = 0, marcoInf = 0, marcoLat = 0;
  let hojaLat = 0, hojaCen = 0, hojaRue = 0;
  const esCompacto = tipoPerfil.toUpperCase().includes("COMPACTO");
  const testero = (!isNaN(T) && T > 0) ? T : 0;

  if (esCompacto) {
    marcoSup = L - 90;
    marcoInf = L - 90;
    marcoLat = H - testero;
    hojaLat = H - testero - 55;
    hojaCen = H - testero - 55;
    hojaRue = (L / 2) - 40;
  } else {
    // Modo estándar 2 Hojas
    marcoSup = L - 37;
    marcoInf = L - 37;
    marcoLat = H;
    hojaLat = H - 55;
    hojaCen = H - 55;
    hojaRue = (L / 2) - 14.5;
  }

  // 4. CÁLCULO DE CRISTALES
  // Ancho_Cristal = Hoja_Ruedas - 56 - Descuento_Cristal
  // Alto_Cristal = Hoja_Lateral - 104 + 28 - Descuento_Cristal
  const anchoCristal = hojaRue - 56 - descCristal;
  const altoCristal = hojaLat - 104 + 28 - descCristal;

  // 5. MEDIO HOJA LATERAL
  const medioHojaLat = hojaLat / 2;

  // 6. ESCRITURA EN LAS COLUMNAS CORRESPONDIENTES
  setVal("Marco_Superior", marcoSup);
  setVal("Marco_Inferior", marcoInf);
  setVal("Marco_Lateral", marcoLat);
  setVal("Hoja_Lateral", hojaLat);
  setVal("Hoja_Central", hojaCen);
  setVal("Hoja_Ruedas", hojaRue);
  setVal("Ancho_Cristal", anchoCristal);
  setVal("Alto_Cristal", altoCristal);
  setVal("Medio_Hoja_Lateral", medioHojaLat);

  return true;
}
