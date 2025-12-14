/**
 * ===============================================
 * ARCHIVO: Reader.gs
 * Funciones de Lectura de Datos (Config, Inventario, Cortes)
 * ===============================================
 */

/**
 * Lee la hoja de 'Configuración' y carga los parámetros globales del sistema.
 * Actualizado para incluir el 'Desperdicio por Corte' en el cálculo de la longitud útil.
 * * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo activa.
 * @returns {Object} Un objeto con todas las claves y valores de configuración.
 * @throws {Error} Si falta la hoja o algún parámetro crítico.
 */
function leerConfiguracion(ss) {
  console.log("[Fase 1] Leyendo Configuración...");
  const hojaConfig = ss.getSheetByName(HOJA_CONFIG_NOMBRE);
  if (!hojaConfig) throw new Error(`No se encuentra la hoja "${HOJA_CONFIG_NOMBRE}"`);
  
  // Leemos todos los datos de la hoja
  const data = hojaConfig.getRange("A2:B" + hojaConfig.getLastRow()).getValues();
  const configObj = {};
  
  // --- PROCESAMIENTO DE PARES CLAVE-VALOR ---
  for (const fila of data) {
    const key = fila[0];
    const value = fila[1];

    // Ignoramos filas vacías o incompletas
    if (key && (value !== null && value !== '')) {
      // Convertimos los espacios del nombre del parámetro a guiones bajos para usar como clave
      const cleanKey = key.toString().replace(/\s+/g, '_');
      
      // Si es un número, lo guardamos como número. Si no, como texto.
      // Esto maneja automáticamente Longitud, Límites, Saneo y el nuevo Desperdicio.
      if (!isNaN(value)) {
        configObj[cleanKey] = Number(value);
      } else {
        configObj[cleanKey] = value;
      }
    }
  }
  
  // --- VALIDACIÓN DE PARÁMETROS CRÍTICOS ---
  // Se añade 'Desperdicio_por_Corte' a la lista de obligatorios.
  const parametrosRequeridos = [
    'Longitud_Barra_Nueva', 
    'Saneo', 
    'Desperdicio_por_Corte', // <-- NUEVO
    'Limite_Desechable', 
    'Limite_Retal', 
    'Limite_Almacenable'
  ];

  const faltantes = parametrosRequeridos.filter(param => configObj[param] === undefined);
  if (faltantes.length > 0) {
    throw new Error(`Faltan valores clave en la hoja 'Configuración': ${faltantes.join(', ')}`);
  }
  
  // --- CÁLCULO DE LA LONGITUD ÚTIL REAL ---
  // La longitud útil es la barra nueva MENOS el saneo de la punta,
  // MENOS el desperdicio del primer corte que se hará.
  // Esto asegura que el motor no intente meter más material del que físicamente cabe.
  configObj.Longitud_Util = configObj.Longitud_Barra_Nueva - configObj.Saneo - configObj.Desperdicio_por_Corte;
  
  console.log(`[Info] Configuración cargada. Longitud de Barra: ${configObj.Longitud_Barra_Nueva}mm. Saneo: ${configObj.Saneo}mm. Desperdicio: ${configObj.Desperdicio_por_Corte}mm. -> Longitud Útil Real: ${configObj.Longitud_Util}mm`);
  
  return configObj;
}

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

// EN ARCHIVO: Reader.gs

/**
 * Función robusta para leer los cortes pendientes de un perfil específico.
 * Incluye filtrado estricto por tipo de ventana (normalizado) y por tipo de hueco (H10/H18).
 */
function leerCortesPendientes(hojaDatos, config) { 
  const ultimaFila = hojaDatos.getLastRow();
  const cortes = [];
  const cortesParaHuella = [];
  // Definimos la columna Tipo aquí por seguridad (asumimos que es la E)
  const COLUMNA_TIPO_LETRA_LOCAL = "E"; 

  if (ultimaFila < FILA_INICIO_DATOS) return { cortes, huella: "" };
  
  // --- LECTURA DE DATOS EN BLOQUE ---
  const rangoAncho = hojaDatos.getRange("B" + FILA_INICIO_DATOS + ":B" + ultimaFila).getValues();
  const rangoAlto = hojaDatos.getRange("C" + FILA_INICIO_DATOS + ":C" + ultimaFila).getValues();
  
  // Columna I: El Checkbox de Control
  const rangoCheck = hojaDatos.getRange(COLUMNA_CALCULADO_LETRA + FILA_INICIO_DATOS + ":" + COLUMNA_CALCULADO_LETRA + ultimaFila).getValues();
  
  // Columna D: Hueco (H10/H18)
  const rangoHueco = hojaDatos.getRange(COLUMNA_HUECO_LETRA + FILA_INICIO_DATOS + ":" + COLUMNA_HUECO_LETRA + ultimaFila).getValues();
  
  // Columna E: Tipo de Ventana
  const rangoTipo = hojaDatos.getRange(COLUMNA_TIPO_LETRA_LOCAL + FILA_INICIO_DATOS + ":" + COLUMNA_TIPO_LETRA_LOCAL + ultimaFila).getValues();
  
  const rangoEtiqueta = hojaDatos.getRange("A" + FILA_INICIO_DATOS + ":A" + ultimaFila).getValues();
  const rangoId = hojaDatos.getRange(COLUMNA_ID_LETRA + FILA_INICIO_DATOS + ":" + COLUMNA_ID_LETRA + ultimaFila).getValues();
  
  // Columna específica de corte del perfil actual
  const rangoCortes = hojaDatos.getRange(config.columna + FILA_INICIO_DATOS + ":" + config.columna + ultimaFila).getValues(); 

  // --- PREPARACIÓN DEL FILTRO DE TIPO (NORMALIZADO Y ROBUSTO) ---
  let tiposValidosNormalizados = null;
  // Usamos 'tipoRequerido' que es la propiedad que definimos en CONFIG.gs
  if (config.tipoRequerido) { 
      // Aseguramos que sea un array. Si es cadena única, la metemos en array.
      const arrayTipos = Array.isArray(config.tipoRequerido) ? config.tipoRequerido : [config.tipoRequerido];
      // Normalizamos a mayúsculas y sin espacios extra para comparación segura
      tiposValidosNormalizados = arrayTipos.map(t => String(t).toUpperCase().trim());
  }

  // --- BUCLE PRINCIPAL DE FILTRADO ---
  for (let i = 0; i < rangoCheck.length; i++) {
    
    // --- 1. LÓGICA INVERSA DEL CHECKBOX ---
    const estaMarcado = rangoCheck[i][0];
    if (estaMarcado === true || String(estaMarcado).toUpperCase() === "TRUE") {
      continue; // Si está marcado (TRUE), ya está hecho -> Saltamos
    }

    // --- 2. VALIDAR MEDIDAS (B y C) ---
    const ancho = parseFloat(String(rangoAncho[i][0]).replace(",", "."));
    const alto = parseFloat(String(rangoAlto[i][0]).replace(",", "."));
    if (isNaN(ancho) || ancho <= 0 || isNaN(alto) || alto <= 0) {
      continue; // Si no tiene medidas válidas -> Saltamos
    }

    // --- 3. NUEVO: FILTRO POR TIPO DE VENTANA (ROBUSTO) ---
    // Si este perfil requiere tipos específicos...
    if (tiposValidosNormalizados) {
        // Leemos y normalizamos el tipo de esta fila
        const tipoFilaNormalizado = String(rangoTipo[i][0]).toUpperCase().trim();
        
        // Si el tipo de la fila NO está en la lista normalizada -> SALTAR
        if (!tiposValidosNormalizados.includes(tipoFilaNormalizado)) {
            continue;
        }
    }

    // --- 4. FILTRO DE PERFIL (H10 vs H18) (COMPARACIÓN ESTRICTA) ---
    // Usamos nombreBase para detectar si es un perfil de hoja H10 o H18
    const nombrePerfilBase = config.nombreBase.toUpperCase();
    // Normalizamos el valor de la celda Hueco: mayúsculas y sin espacios
    const huecoFila = String(rangoHueco[i][0]).toUpperCase().trim();

    if (nombrePerfilBase.includes("H10")) {
        // Si el perfil es H10, la fila TIENE QUE SER exactamente "H10"
        if (huecoFila !== "H10") {
            continue; // Si es "H18" o cualquier otra cosa, salta
        }
    }
    else if (nombrePerfilBase.includes("H18")) {
        // Si el perfil es H18, la fila TIENE QUE SER exactamente "H18"
        if (huecoFila !== "H18") {
            continue; // Si es "H10" o cualquier otra cosa, salta
        }
    }
    // Si no es un perfil de hoja, este bloque se ignora y pasa.

    // --- 5. RECOGER DATOS DEL CORTE ---
    const valorCorte = parseFloat(String(rangoCortes[i][0]).replace(",", "."));
    const valorId = rangoId[i][0]; 
    const valorEtiqueta = rangoEtiqueta[i][0];

    if (valorCorte > 0 && valorId) {
      // Añadimos tantos cortes como diga el multiplicador
      for (let j = 0; j < config.multiplicador; j++) {
        cortes.push({ 
          id: valorId, 
          etiqueta: valorEtiqueta, 
          medida: valorCorte,
          perfil: config.nombre // Usamos el nombre específico (ej. "Marco Lateral Guía")
        });
        cortesParaHuella.push(valorCorte);
      }
    }
  }

  // --- FINALIZACIÓN ---
  // Ordenamos los cortes de mayor a menor para el algoritmo Greedy
  cortes.sort((a, b) => b.medida - a.medida);
  cortesParaHuella.sort((a, b) => a - b);

  return {
    cortes: cortes,
    huella: cortesParaHuella.join(',')
  };
}
