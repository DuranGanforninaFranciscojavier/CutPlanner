/**
 * ===============================================
 * ARCHIVO: Reader.gs
 * Funciones de Lectura de Datos (Config, Inventario, Cortes)
 * Adaptado a Mapeo Dinámico de Encabezados (Fila 1)
 * ===============================================
 */

/**
 * Lee la hoja de 'Configuración' y carga los parámetros globales del sistema.
 * Normaliza las claves a formato snake_case (ej. 'Descuento_ Ventana' -> 'Descuento_Ventana').
 */
function leerConfiguracion(ss) {
  console.log("[Fase 1] Leyendo Configuración...");
  const hojaConfig = ss.getSheetByName(HOJA_CONFIG_NOMBRE);
  if (!hojaConfig) throw new Error(`No se encuentra la hoja "${HOJA_CONFIG_NOMBRE}"`);
  
  const data = hojaConfig.getRange("A2:B" + hojaConfig.getLastRow()).getValues();
  const configObj = {};
  
  for (const fila of data) {
    const key = fila[0];
    const value = fila[1];

    if (key && (value !== null && value !== '')) {
      // Normalizamos la clave colapsando espacios y guiones a un único guion bajo
      const cleanKey = key.toString().trim().replace(/[\s_-]+/g, '_');
      
      if (!isNaN(value) && value !== '') {
        configObj[cleanKey] = Number(value);
      } else {
        configObj[cleanKey] = value;
      }
    }
  }
  
  // Parámetros críticos requeridos para el motor
  const parametrosRequeridos = [
    'Longitud_Barra_Nueva', 
    'Saneo', 
    'Desperdicio_por_Corte',
    'Limite_Desechable', 
    'Limite_Retal', 
    'Limite_Almacenable'
  ];

  const faltantes = parametrosRequeridos.filter(param => configObj[param] === undefined);
  if (faltantes.length > 0) {
    throw new Error(`Faltan valores clave en la hoja 'Configuración': ${faltantes.join(', ')}`);
  }
  
  configObj.Longitud_Util = configObj.Longitud_Barra_Nueva - configObj.Saneo - configObj.Desperdicio_por_Corte;
  console.log(`[Info] Configuración cargada. Longitud Útil Real: ${configObj.Longitud_Util}mm`);
  
  return configObj;
}

/**
 * Lee el inventario físico disponible en la hoja 'Inventario'.
 */
function leerInventario(ss) {
  console.log("[Fase 1] Leyendo Inventario...");
  const hojaInv = ss.getSheetByName(HOJA_INVENTARIO_NOMBRE);
  if (!hojaInv) throw new Error(`No se encuentra la hoja "${HOJA_INVENTARIO_NOMBRE}"`);

  const inventario = [];
  if (hojaInv.getLastRow() < 2) {
    console.warn("[Aviso] La hoja de Inventario está vacía.");
    return inventario;
  }
  
  const data = hojaInv.getRange(2, 1, hojaInv.getLastRow() - 1, 8).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const fila = data[i];
    const filaNum = i + 2;
    const id = fila[0];       
    const serie = fila[1];    
    const perfil = fila[2];   
    const tipo = fila[3];     
    const etiqueta = fila[6]; 
    const longitud = parseFloat(fila[4]); 
    const cantidad = parseInt(fila[5], 10); 
    
    if (!isNaN(cantidad) && cantidad > 0) { 
      for (let j = 0; j < cantidad; j++) {
        inventario.push({
          filaInventario: filaNum,
          idMaterial: id,
          etiquetaMaterial: etiqueta,
          perfil: perfil,
          longitud: longitud,
          tipo: tipo,
          esVirtual: false
        });
      }
    }
  }
  console.log(`[Info] Inventario cargado. ${inventario.length} barras disponibles encontradas.`);
  return inventario;
}

/**
 * Lee los cortes pendientes de un perfil específico usando mapeo dinámico de encabezados.
 */
function leerCortesPendientes(hojaDatos, config) { 
  const ultimaFila = hojaDatos.getLastRow();
  const cortes = [];
  const cortesParaHuella = [];

  if (ultimaFila < FILA_INICIO_DATOS) return { cortes, huella: "" };

  const colMap = getHeaderMap(hojaDatos, FILA_CABECERA);

  // Mapeo dinámico de columnas necesarias
  const colCheck = colMap.getCol("Estado_Ventana") || colMap.getCol("Calculado") || colMap.getCol("Hecha");
  const colId = colMap.getCol("Id_Ventana") || colMap.getCol("ID");
  const colEtiqueta = colMap.getCol("Etiqueta");
  const colAncho = colMap.getCol("Ancho_Corte") || colMap.getCol("Ancho_Luz");
  const colAlto = colMap.getCol("Alto_Corte") || colMap.getCol("Alto_Luz");
  const colHueco = colMap.getCol("Hueco_Cristal") || colMap.getCol("Hueco");
  const colTipo = colMap.getCol("Tipo_Perfil") || colMap.getCol("Tipo");
  const colCorte = colMap.getCol(config.columna) || colMap.getCol(config.nombre);

  if (!colCheck || !colId || !colCorte) {
    console.warn(`[Reader] No se pudieron localizar todas las columnas necesarias para ${config.nombre}`);
    return { cortes, huella: "" };
  }

  // Lectura en bloque de todas las columnas de datos
  const numCols = hojaDatos.getLastColumn();
  const filasDatos = hojaDatos.getRange(FILA_INICIO_DATOS, 1, ultimaFila - FILA_INICIO_DATOS + 1, numCols).getValues();

  // Normalización de filtro por tipo de ventana
  let tiposValidosNormalizados = null;
  if (config.tipoRequerido) {
    const arrayTipos = Array.isArray(config.tipoRequerido) ? config.tipoRequerido : [config.tipoRequerido];
    tiposValidosNormalizados = arrayTipos.map(t => String(t).toUpperCase().replace(/[\s_-]+/g, ''));
  }

  for (let i = 0; i < filasDatos.length; i++) {
    const fila = filasDatos[i];

    // 1. Checkbox: Si está marcado (TRUE), se ignora
    const estaMarcado = fila[colCheck - 1];
    if (estaMarcado === true || String(estaMarcado).toUpperCase() === "TRUE") {
      continue;
    }

    // 2. Validación de dimensiones
    const ancho = parseFloat(String(fila[colAncho - 1] || "").replace(",", "."));
    const alto = parseFloat(String(fila[colAlto - 1] || "").replace(",", "."));
    if (isNaN(ancho) || ancho <= 0 || isNaN(alto) || alto <= 0) {
      continue;
    }

    // 3. Filtro por tipo de ventana (Compacto / 2 Hojas)
    if (tiposValidosNormalizados && colTipo) {
      const tipoFila = String(fila[colTipo - 1] || "").toUpperCase().replace(/[\s_-]+/g, '');
      if (!tiposValidosNormalizados.includes(tipoFila)) {
        continue;
      }
    }

    // 4. Filtro por tipo de hueco (H10 vs H18)
    const nombrePerfilBase = String(config.nombreBase || config.nombre).toUpperCase();
    const huecoFila = String(fila[colHueco - 1] || "").toUpperCase().trim();

    if (nombrePerfilBase.includes("H10") && huecoFila !== "H10") {
      continue;
    }
    if (nombrePerfilBase.includes("H18") && huecoFila !== "H18") {
      continue;
    }

    // 5. Recoger datos de corte
    const valorCorte = parseFloat(String(fila[colCorte - 1] || "").replace(",", "."));
    const valorId = fila[colId - 1]; 
    const valorEtiqueta = colEtiqueta ? fila[colEtiqueta - 1] : `V-${i + 1}`;

    if (valorCorte > 0 && valorId) {
      for (let j = 0; j < config.multiplicador; j++) {
        cortes.push({ 
          id: valorId, 
          etiqueta: valorEtiqueta, 
          medida: valorCorte,
          perfil: config.nombre 
        });
        cortesParaHuella.push(valorCorte);
      }
    }
  }

  // Ordenar de mayor a menor para optimización Greedy
  cortes.sort((a, b) => b.medida - a.medida);
  cortesParaHuella.sort((a, b) => a - b);

  return {
    cortes: cortes,
    huella: cortesParaHuella.join(',')
  };
}
