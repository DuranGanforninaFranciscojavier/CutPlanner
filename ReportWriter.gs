/**
 * ===============================================
 * ARCHIVO: ReportWriter.gs
 * Escritura y Formateo del Informe Histórico
 * ===============================================
 */

function prepararHojaHistorico(ss) {
  let hoja = ss.getSheetByName(HOJA_HISTORICO_NOMBRE);
  const numColumnas = CABECERAS_HISTORICO.length; 

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_HISTORICO_NOMBRE);
    hoja.appendRow(CABECERAS_HISTORICO);
    hoja.setFrozenRows(1);
    console.log(`[Info] Hoja "${HOJA_HISTORICO_NOMBRE}" creada.`);
    aplicarFormatoBase(hoja); 
  } else {
    if (hoja.getRange("D1").getValue() !== "Etiqueta Mat") { 
      hoja.insertRowBefore(1);
      hoja.deleteRow(2); 
      hoja.getRange(1, 1, 1, numColumnas).setValues([CABECERAS_HISTORICO]);
      hoja.setFrozenRows(1);
      console.log(`[Info] Cabeceras actualizadas a v37.`);
      aplicarFormatoBase(hoja); 
    }
  }
  return hoja;
}

function aplicarFormatoBase(hoja) {
  const maxFilas = hoja.getMaxRows();
  const numColumnas = CABECERAS_HISTORICO.length; 

  const rangoCompleto = hoja.getRange(1, 1, maxFilas, numColumnas);
  rangoCompleto.setVerticalAlignment('top').setHorizontalAlignment('left');

  hoja.getRange(1, 1, 1, numColumnas).setBackground(COLOR_FONDO_CABECERA).setFontColor(COLOR_TEXTO_CABECERA);
  hoja.getRange(2, 1, maxFilas - 1, numColumnas).setFontColor(null); 

  hoja.getRange(1, COL_CORTES_IDX, maxFilas, 1).setWrap(true);
  hoja.getRange(1, COL_ID_IDX, maxFilas, 1).setFontColor(COLOR_TEXTO_GRIS);
  hoja.getRange(1, COL_FECHA_IDX, maxFilas, 1).setFontColor(COLOR_TEXTO_GRIS);
  
  const rangoHuella = hoja.getRange(1, COL_HUELLA_IDX, maxFilas, 1);
  rangoHuella.setFontColor(COLOR_TEXTO_GRIS);
  rangoHuella.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  hoja.getRange(1, 1, 1, numColumnas).setFontColor(COLOR_TEXTO_CABECERA);
}

function cargarHuellasExistentes(hoja) {
  const huellas = new Set();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return huellas; 
  const rangoHuellas = hoja.getRange(2, COL_HUELLA_IDX, ultimaFila - 1, 1).getValues();
  for (const fila of rangoHuellas) {
    if (fila[0]) huellas.add(fila[0]);
  }
  return huellas;
}

/**
 * Escribe los resultados de la optimización en la hoja de histórico 'Plan_Cortes'.
 * (VERSIÓN ACTUALIZADA CON DESPERDICIO DE CORTE)
 * Mantiene toda la lógica de formato, colores y generación de ID original.
 */
function escribirResultadosEnHistorico(hoja, solucion, nombrePerfil, piezasTotales, huella) {
  // ... (Inicio de la función: cálculo de ID, fechas, logs... IGUAL QUE ANTES) ...
  const configPerfil = TODOS_PERFILES.find(p => p.nombre === nombrePerfil);
  const siglasId = configPerfil ? configPerfil.id.split('-')[1] : "GEN"; 

  const idOptimizacion = `${siglasId}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss")}`;
  const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy HH:mm");

  console.log(`[Fase 3] Escribiendo ${solucion.length} filas para el ID: ${idOptimizacion}`);

  const filasNuevas = [];
  const filasParaFormato = [];
  const nuevosRestos = [];
  const materialUsado = [];
  const desperdicioPorCorte = CONFIG.Desperdicio_por_Corte || 4;

  for (let i = 0; i < solucion.length; i++) {
    const barra = solucion[i];
    // ... (Cálculo del restoRealFinal IGUAL QUE ANTES) ...
    const sumaLongitudesNetas = barra.cortes.reduce((sum, c) => sum + c.medida, 0);
    const numeroDeCortes = barra.cortes.length;
    const desperdicioTotalBarra = numeroDeCortes * desperdicioPorCorte;
    const longitudTotalOcupada = sumaLongitudesNetas + desperdicioTotalBarra;
    let restoRealFinal = barra.longitud - longitudTotalOcupada;
    if (barra.tipo === 'Barra Nueva' || barra.esVirtual) {
       restoRealFinal -= (CONFIG.Saneo || 0);
    }
    const resto = restoRealFinal;
    // -----------------------------------------------------

    materialUsado.push(barra);

    // Determinación del COLOR y creación de objetos de inventario
    // Ya no necesitamos la variable de texto 'tipoDeResto' para la hoja
    let colorResto = null;
    let nuevoRestoObjeto = null;
    
    if (resto < CONFIG.Limite_Desechable) { 
      // Desechable (Rojo claro)
      colorResto = COLOR_RESTO_DESECHABLE;
    } else if (resto >= CONFIG.Limite_Almacenable) { 
      // Almacenable (Azul claro)
      colorResto = COLOR_RESTO_ALMACENABLE;
      nuevoRestoObjeto = { perfil: nombrePerfil, longitud: resto, tipo: 'Almacenable' }; 
    } else if (resto >= CONFIG.Limite_Retal) { 
      // Retal (Verde claro)
      colorResto = COLOR_RESTO_RETAL;
      nuevoRestoObjeto = { perfil: nombrePerfil, longitud: resto, tipo: 'Retal' }; 
    }
    
    if (nuevoRestoObjeto) nuevosRestos.push(nuevoRestoObjeto);
    
    const cortesArray = barra.cortes.map(corte => `- ${corte.medida} (${corte.etiqueta})`);
    const cortesString = cortesArray.join('\n');
    
    // --- CAMBIO 1: Array de datos más corto ---
    filasNuevas.push([
      idOptimizacion,
      "GP60",
      nombrePerfil,
      barra.esVirtual ? `${barra.etiquetaMaterial}` : barra.etiquetaMaterial, 
      barra.tipo,                                                         
      barra.longitud,
      cortesString,
      Math.round(resto),
      false,
      fechaStr,
      piezasTotales,
      // REMOVIDO: tipoDeResto,
      huella
    ]);
    
    // Guardamos los colores para aplicarlos después
    filasParaFormato.push({
      index: hoja.getLastRow() + filasNuevas.length,
      colorResto: colorResto, // El color determinado (Rojo/Verde/Azul)
      colorPerfil: TODOS_PERFILES.find(p => p.nombre === nombrePerfil)?.color || '#ffffff',
      colorTipo: COLORES_TIPO_HISTORICO[barra.tipo] || '#ffffff' 
    });
  }

  // Escritura y formateo en bloque
  if (filasNuevas.length > 0) {
    const filaInicio = hoja.getLastRow() + 1;
    const rangoEscritura = hoja.getRange(filaInicio, 1, filasNuevas.length, CABECERAS_HISTORICO.length);
    rangoEscritura.setValues(filasNuevas);
    rangoEscritura.setBackground(COLOR_FONDO_BASE);
    
    for (const filaInfo of filasParaFormato) {
      const filaActual = filaInfo.index;
      hoja.getRange(filaActual, COL_PERFIL_IDX).setBackground(filaInfo.colorPerfil); 
      hoja.getRange(filaActual, COL_TIPO_MAT_IDX).setBackground(filaInfo.colorTipo); 
      
      // --- CAMBIO 2: Aplicar color a la columna RESTO ---
      if (filaInfo.colorResto) {
        // Usamos COL_RESTO_IDX (Columna 8/H) en lugar de la eliminada
        hoja.getRange(filaActual, COL_RESTO_IDX).setBackground(filaInfo.colorResto); 
      }
      hoja.getRange(filaActual, COL_HECHA_IDX).insertCheckboxes(); 
    }
  }
  
  return [nuevosRestos, materialUsado];
}

/**
 * Escribe la lista de barras necesarias
 */
function generarInformeMaterial(ss, mapaResumen) {
  const hoja = ss.getSheetByName(HOJA_RESUMEN_NOMBRE);
  const filas = [];
  
  mapaResumen.forEach((cantidad, nombrePerfil) => {
    // Buscamos el color del perfil para ponerlo bonito
    const infoPerfil = TODOS_PERFILES.find(p => p.nombre === nombrePerfil);
    const color = infoPerfil ? infoPerfil.color : '#ffffff';
    
    filas.push([nombrePerfil, cantidad, ""]); // La col C la usaremos para pintar el fondo
  });

  if (filas.length > 0) {
    const rango = hoja.getRange(2, 1, filas.length, 3);
    
    // Escribir datos (Solo cols A y B tienen texto relevante)
    hoja.getRange(2, 1, filas.length, 2).setValues(filas.map(f => [f[0], f[1]]));
    
    // Aplicar colores visuales
    filas.forEach((fila, index) => {
      const nombre = fila[0];
      const infoPerfil = TODOS_PERFILES.find(p => p.nombre === nombre);
      if (infoPerfil) {
        // Pintamos toda la fila del color del perfil para identificar rápido
        hoja.getRange(index + 2, 1, 1, 3).setBackground(infoPerfil.color);
      }
    });
    
    // Ajustar anchos
    hoja.autoResizeColumns(1, 3);
  } else {
    hoja.getRange("A2").setValue("No se necesita material nuevo.");
  }
}