/**
 * ===============================================
 * ARCHIVO: CutQueueManager.gs
 * Gestión de la Cola de Cortes y Planes de Barra
 * Estructura normalizada pieza a pieza (AppSheet Ready)
 * ===============================================
 */

/**
 * Prepara la hoja 'Plan_Cortes' con la estructura tabular requerida por AppSheet.
 * Garantiza que no haya celdas combinadas ni formatos que rompan la base de datos.
 */
function prepararHojaHistorico(ss) {
  let hoja = ss.getSheetByName(HOJA_HISTORICO_NOMBRE);
  const numColumnas = CABECERAS_HISTORICO.length;

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_HISTORICO_NOMBRE);
    hoja.appendRow(CABECERAS_HISTORICO);
    hoja.setFrozenRows(1);
    aplicarFormatoBase(hoja);
    console.log(`[Info] Hoja "${HOJA_HISTORICO_NOMBRE}" creada con esquema AppSheet.`);
  } else {
    const primeraCabecera = hoja.getRange("A1").getValue();
    if (primeraCabecera !== CABECERAS_HISTORICO[0]) {
      hoja.insertRowBefore(1);
      hoja.deleteRow(2);
      hoja.getRange(1, 1, 1, numColumnas).setValues([CABECERAS_HISTORICO]);
      hoja.setFrozenRows(1);
      aplicarFormatoBase(hoja);
      console.log(`[Info] Cabeceras de "${HOJA_HISTORICO_NOMBRE}" actualizadas.`);
    }
  }
  return hoja;
}

/**
 * Aplica el formato base a la cabecera (sin combinar celdas).
 */
function aplicarFormatoBase(hoja) {
  const numColumnas = CABECERAS_HISTORICO.length;
  const rangoCabecera = hoja.getRange(1, 1, 1, numColumnas);
  rangoCabecera.setBackground(COLOR_FONDO_CABECERA)
               .setFontColor(COLOR_TEXTO_CABECERA)
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
  hoja.setFrozenRows(1);
}

/**
 * Obtiene el conjunto de huellas existentes para evitar optimizaciones duplicadas.
 */
function cargarHuellasExistentes(hoja) {
  const huellas = new Set();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return huellas;

  const rangoHuellas = hoja.getRange(2, COL_HUELLA_IDX, ultimaFila - 1, 1).getValues();
  for (const fila of rangoHuellas) {
    if (fila[0]) huellas.add(String(fila[0]).trim());
  }
  return huellas;
}

/**
 * Registra los resultados de la optimización desglosados PIEZA A PIEZA.
 * Cada fila representa un corte individual vinculado a su barra física.
 * Retorna [nuevosRestos, materialUsado] para que InventoryManager actualice el stock.
 */
function escribirResultadosEnHistorico(hoja, solucion, nombrePerfil, piezasTotales, huella, nombreProyecto) {
  const configPerfil = TODOS_PERFILES.find(p => p.nombre === nombrePerfil);
  const siglasId = configPerfil ? configPerfil.id.split('-')[1] : "GEN";
  const timestampLote = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMMdd-HHmmss");
  const fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const idOptimizacion = `${siglasId}-${timestampLote}`;
  const proyecto = nombreProyecto || "PROYECTO GENERAL";

  console.log(`[CutQueueManager] Registrando ${solucion.length} barras optimizadas para: ${nombrePerfil}`);

  const filasNuevas = [];
  const nuevosRestos = [];
  const materialUsado = [];
  const desperdicioPorCorte = CONFIG.Desperdicio_por_Corte || 4;

  for (let i = 0; i < solucion.length; i++) {
    const barra = solucion[i];
    const idBarra = `${idOptimizacion}-B${i + 1}`;

    // Cálculo del resto físico final
    const sumaLongitudesNetas = barra.cortes.reduce((sum, c) => sum + c.medida, 0);
    const numeroDeCortes = barra.cortes.length;
    const desperdicioTotalBarra = numeroDeCortes * desperdicioPorCorte;
    const longitudTotalOcupada = sumaLongitudesNetas + desperdicioTotalBarra;
    let restoRealFinal = barra.longitud - longitudTotalOcupada;

    if (barra.tipo === 'Barra Nueva' || barra.esVirtual) {
      restoRealFinal -= (CONFIG.Saneo || 0);
    }

    const restoRedondeado = Math.round(restoRealFinal);

    // Clasificación del resto
    let tipoDeResto = "DESECHO";
    let nuevoRestoObjeto = null;

    if (restoRedondeado < CONFIG.Limite_Desechable) {
      tipoDeResto = "DESECHO";
    } else if (restoRedondeado >= CONFIG.Limite_Almacenable) {
      tipoDeResto = "ALMACENABLE";
      nuevoRestoObjeto = { perfil: nombrePerfil, longitud: restoRealFinal, tipo: 'Almacenable' };
    } else if (restoRedondeado >= CONFIG.Limite_Retal) {
      tipoDeResto = "RETAL";
      nuevoRestoObjeto = { perfil: nombrePerfil, longitud: restoRealFinal, tipo: 'Retal' };
    }

    if (nuevoRestoObjeto) {
      nuevosRestos.push(nuevoRestoObjeto);
    }

    materialUsado.push(barra);

    // Desglose ATÓMICO: 1 fila por cada corte/pieza de la barra
    for (let k = 0; k < barra.cortes.length; k++) {
      const corte = barra.cortes[k];
      const posicionCorte = k + 1;
      const idCorte = `${idBarra}-C${posicionCorte}`;

      filasNuevas.push([
        idCorte,                                             // ID_Corte (Key único)
        idBarra,                                             // ID_Barra (Ref)
        proyecto,                                            // Proyecto
        corte.etiqueta || "V-?",                             // Ventana
        nombrePerfil,                                        // Perfil
        corte.medida,                                        // Medida_Corte
        posicionCorte,                                       // Posición de corte en la barra
        barra.tipo,                                          // Tipo_Material (Barra Nueva / Retal)
        barra.etiquetaMaterial || "",                        // Etiqueta_Material
        barra.longitud,                                      // Long_Barra
        restoRedondeado,                                     // Resto_Barra
        tipoDeResto,                                         // Tipo_Resto (DESECHO/RETAL/ALMACENABLE)
        "Optimizado",                                        // Estado
        fechaStr,                                            // Fecha_Hora
        huella                                               // Huella de cortes
      ]);
    }
  }

  // Escritura en bloque en la hoja
  if (filasNuevas.length > 0) {
    const filaInicio = hoja.getLastRow() + 1;
    const rangoEscritura = hoja.getRange(filaInicio, 1, filasNuevas.length, CABECERAS_HISTORICO.length);
    rangoEscritura.setValues(filasNuevas);
    rangoEscritura.setHorizontalAlignment("center");
  }

  return [nuevosRestos, materialUsado];
}

/**
 * Genera el resumen consolidado de material nuevo a comprar en la hoja 'Material_Necesario'.
 */
function generarInformeMaterial(ss, mapaResumen) {
  const hoja = ss.getSheetByName(HOJA_RESUMEN_NOMBRE);
  if (!hoja) return;

  const filas = [];
  mapaResumen.forEach((cantidad, nombrePerfil) => {
    filas.push([nombrePerfil, cantidad]);
  });

  const lastRow = hoja.getLastRow();
  if (lastRow > 1) {
    hoja.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, 2).setValues(filas);
    hoja.autoResizeColumns(1, 2);
  } else {
    hoja.getRange("A2").setValue("No se necesita material nuevo.");
  }
}
