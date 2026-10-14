/**
 * ===============================================
 * ARCHIVO: Main.gs
 * Punto de entrada y Orquestación Principal
 * ===============================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🏭 Gestión de Inventario')
      .addItem('▶️ PRODUCCIÓN (Usa Stock y Actualiza)', 'ejecutarOptimizacion') 
      .addSeparator()
      .addItem('🧮 MODO SOLO (Solo Calcular)', 'ejecutarModoSolo')
      .addSeparator()
      .addItem('🔄 Recalcular Todo', 'actualizarTodo')
      .addToUi();
}

/**
 * ==============================================================================
 * MODO SOLO / CALCULADORA
 * - Calcula necesidades sin usar inventario real.
 * - Usa presupuesto de tiempo dinámico y holgura de stock virtual.
 * ==============================================================================
 */
function ejecutarModoSolo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const horaInicioTotal = new Date();
  ALERTAS = []; 
  
  let resumenCompra = new Map(); 
  let todosLosIdsParaMarcar = []; 
  let todosLosRetalesGenerados = [];
  let trabajosProcesados = 0;

  try {
    console.log("--- 🧮 INICIO MODO SOLO (PRODUCCIÓN SIN STOCK) ---");
    
    // 1. SETUP COMÚN
    CONFIG = leerConfiguracion(ss);
    const hojaDatos = ss.getSheetByName(HOJA_DATOS_NOMBRE);
    const hojaHistorico = prepararHojaHistorico(ss);
    
    // 2. OBTENER PROYECTO ACTIVO
    const nombreProyecto = obtenerNombreProyecto(hojaDatos);

    // 3. ANÁLISIS Y ORDENACIÓN (Específico de Modo Solo)
    console.log("--- 📋 Analizando complejidad de tareas... ---");
    let tareasPendientes = [];
    const longitudBarraUtil = CONFIG.Longitud_Util;

    for (const configPerfil of TODOS_PERFILES) {
      const trabajoPrevio = leerCortesPendientes(hojaDatos, configPerfil);
      const numCortes = trabajoPrevio.cortes.length;
      if (numCortes > 0) {
        const longitudTotalNecesaria = trabajoPrevio.cortes.reduce((sum, c) => sum + c.medida, 0);
        tareasPendientes.push({
          config: configPerfil,
          numCortes: numCortes,
          longitudTotal: longitudTotalNecesaria
        });
      }
    }
    // Ordenar por complejidad
    tareasPendientes.sort((a, b) => a.numCortes - b.numCortes);

    if (tareasPendientes.length === 0) {
       finalizarEjecucionComun(ss, 0, [], horaInicioTotal, "Modo Solo");
       return;
    }

    // 4. CÁLCULO DEL PRESUPUESTO DE TIEMPO
    const TIEMPO_TOTAL_DISPONIBLE_SEG = 300; 
    let pesoTotal = 0;
    tareasPendientes.forEach(t => { t.peso = t.numCortes * t.numCortes; pesoTotal += t.peso; });
    
    tareasPendientes.forEach(t => {
      t.tiempoAsignado = Math.max(15, Math.round((t.peso / pesoTotal) * TIEMPO_TOTAL_DISPONIBLE_SEG));
    });

    // 5. BUCLE PRINCIPAL
    for (const tarea of tareasPendientes) { 
      const configPerfil = tarea.config;
      const nombrePerfil = configPerfil.nombre;
      const trabajo = leerCortesPendientes(hojaDatos, configPerfil); 
      
      console.log(`\n▶️ PROCESANDO: ${nombrePerfil} (${trabajo.cortes.length} cortes) ---`);
      todosLosIdsParaMarcar.push(...trabajo.cortes.map(c => c.id));

      // Stock Virtual con Holgura (x2)
      let barrasMinimasTeoricas = Math.ceil(tarea.longitudTotal / longitudBarraUtil);
      let barrasParaIntentar = barrasMinimasTeoricas * 2;
      let stockVirtual = [];
      for (let i = 1; i <= barrasParaIntentar; i++) {
        stockVirtual.push(crearBarraVirtual(nombrePerfil, i));
      }

      let solucionFinal = findSolution_v31(trabajo.cortes, stockVirtual, tarea.tiempoAsignado);

      if (solucionFinal) {
        let barrasUsadasReales = solucionFinal.filter(b => b.esVirtual).length;
        if (barrasUsadasReales > 0) resumenCompra.set(nombrePerfil, barrasUsadasReales);
        
        const [nuevosRetales, materialUsado] = escribirResultadosEnHistorico(hojaHistorico, solucionFinal, nombrePerfil, trabajo.cortes.length, "MODO_SOLO", nombreProyecto);
        if (nuevosRetales && nuevosRetales.length > 0) todosLosRetalesGenerados.push(...nuevosRetales);
        
        trabajosProcesados++;
      } else {
        console.error(`🛑 [ABORTADO] Fallo en ${nombrePerfil}`);
        ALERTAS.push(`❌ ${nombrePerfil}: Fallo de optimización complejo.`);
      }
    } 

    // 6. FINALIZACIÓN ESPECÍFICA (Informe de Material)
    console.log("--- Generando Informe Final de Material ---");
    generarInformeMaterial(ss, resumenCompra, todosLosRetalesGenerados);
    hojaHistorico.autoResizeColumns(1, CABECERAS_HISTORICO.length);

    // 7. FINALIZACIÓN COMÚN (Vistas, Alertas, Marcado)
    finalizarEjecucionComun(ss, trabajosProcesados, todosLosIdsParaMarcar, horaInicioTotal, "Modo Solo");

  } catch (e) {
    console.error(e);
    SpreadsheetApp.getUi().alert("❌ ERROR CRÍTICO MODO SOLO:\n" + e.message);
  }
}

/**
 * ==============================================================================
 * MODO PRODUCCIÓN / INVENTARIO
 * - Usa el stock real de la hoja 'Inventario'.
 * - Actualiza el inventario al terminar.
 * ==============================================================================
 */
function ejecutarOptimizacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const horaInicioTotal = new Date();
  ALERTAS = []; 
  let todosLosIdsParaMarcar = []; 
  let trabajosProcesados = 0;
  const TIEMPO_LIMITE_POR_INTENTO_SEG = 120; 

  try {
    console.log("--- 🚀 INICIO OPTIMIZACIÓN (MODO INVENTARIO) ---");
    
    // 1. SETUP COMÚN
    CONFIG = leerConfiguracion(ss);
    const inventarioCompleto = leerInventario(ss);
    const configSheet = ss.getSheetByName(HOJA_CONFIG_NOMBRE);
    const inventarioSheet = ss.getSheetByName(HOJA_INVENTARIO_NOMBRE);
    const hojaDatos = ss.getSheetByName(HOJA_DATOS_NOMBRE);
    
    if (!configSheet || !inventarioSheet || !hojaDatos) throw new Error("Faltan hojas esenciales.");
    
    const hojaHistorico = prepararHojaHistorico(ss);
    const huellasExistentes = cargarHuellasExistentes(hojaHistorico);

    // 2. OBTENER PROYECTO ACTIVO
    const nombreProyecto = obtenerNombreProyecto(hojaDatos);

    // 3. BUCLE PRINCIPAL
    for (const configPerfil of TODOS_PERFILES) { 
      const nombrePerfil = configPerfil.nombre;
      const trabajo = leerCortesPendientes(hojaDatos, configPerfil);
      
      if (trabajo.cortes.length === 0) continue;
      console.log(`--- 🔄 Procesando: ${nombrePerfil} (${trabajo.cortes.length} cortes) ---`);
      
      if (huellasExistentes.has(trabajo.huella)) {
        ALERTAS.push(`--- ${nombrePerfil} ---\nℹ️ Plan saltado (ya existe en histórico).`);
        continue; 
      }

      ALERTAS.push(`--- ${nombrePerfil} ---`);
      
      let stockDisponible = inventarioCompleto.filter(item => item.perfil === nombrePerfil);
      stockDisponible = sortStock(stockDisponible); 

      let barrasNuevasVirtuales = 0;
      let solucionFinal = null;
      let stockParaIntentar = [...stockDisponible]; 
      const LIMITE_BARRAS_VIRTUALES = trabajo.cortes.length + 2;

      while (solucionFinal == null && barrasNuevasVirtuales <= LIMITE_BARRAS_VIRTUALES) {
        horaInicioAlgoritmo = new Date(); 
        let solucionIntento = findSolution_v31(trabajo.cortes, stockParaIntentar, TIEMPO_LIMITE_POR_INTENTO_SEG);
        
        if (solucionIntento) {
          solucionFinal = solucionIntento;
        } else {
          barrasNuevasVirtuales++;
          if (barrasNuevasVirtuales > LIMITE_BARRAS_VIRTUALES) {
              ALERTAS.push(`❌ ${nombrePerfil}: No se encontró solución.`);
              break;
          }
          stockParaIntentar.push(crearBarraVirtual(nombrePerfil, barrasNuevasVirtuales));
        }
      } 

      if (solucionFinal) {
        if (barrasNuevasVirtuales > 0) ALERTAS.push(`❗️ Se necesitaron ${barrasNuevasVirtuales} BARRAS NUEVAS.`);
        
        const [nuevosRetales, materialUsado] = escribirResultadosEnHistorico(hojaHistorico, solucionFinal, nombrePerfil, trabajo.cortes.length, trabajo.huella, nombreProyecto);
        actualizarInventario(ss, materialUsado, nuevosRetales);
        
        const alertaStock = comprobarStockMinimo(ss, nombrePerfil);
        if (alertaStock) ALERTAS.push(alertaStock);
        
        todosLosIdsParaMarcar.push(...trabajo.cortes.map(c => c.id)); 
        trabajosProcesados++;
      }
    } 
    
    // 4. FINALIZACIÓN ESPECÍFICA
    SpreadsheetApp.flush();
    if (typeof aplicarFormatosInventario === 'function') {
      aplicarFormatosInventario(configSheet, inventarioSheet); 
      SpreadsheetApp.flush(); 
    } 

    // 5. FINALIZACIÓN COMÚN
    finalizarEjecucionComun(ss, trabajosProcesados, todosLosIdsParaMarcar, horaInicioTotal, "Modo Producción");

  } catch (e) {
    console.error(e);
    SpreadsheetApp.getUi().alert("❌ ERROR MODO INVENTARIO:\n" + e.message);
  }
}

/**
 * ==============================================================================
 * FUNCIONES AUXILIARES DE ORQUESTACIÓN
 * ==============================================================================
 */

/**
 * Detecta si hay trabajos pendientes y extrae el Nombre del Proyecto usando mapeo dinámico.
 */
function obtenerNombreProyecto(hojaDatos) {
  try {
    const lastRow = hojaDatos.getLastRow();
    if (lastRow < FILA_INICIO_DATOS) return "PROYECTO GENERAL";

    const colMap = getHeaderMap(hojaDatos, FILA_CABECERA);
    const colProyecto = colMap.getCol("Nombre_Proyecto") || colMap.getCol("Proyecto");
    const colCheck = colMap.getCol("Estado_Ventana") || colMap.getCol("Calculado") || colMap.getCol("Hecha");

    if (colProyecto && colCheck) {
      const numFilas = lastRow - FILA_INICIO_DATOS + 1;
      const rangoChecks = hojaDatos.getRange(FILA_INICIO_DATOS, colCheck, numFilas, 1).getValues();
      const rangoNombres = hojaDatos.getRange(FILA_INICIO_DATOS, colProyecto, numFilas, 1).getValues();

      for (let i = 0; i < rangoChecks.length; i++) {
        const hecha = rangoChecks[i][0];
        if (hecha !== true && String(hecha).toUpperCase() !== "TRUE") {
          const nombre = String(rangoNombres[i][0] || "").trim();
          if (nombre) return nombre;
        }
      }
    }
  } catch(e) {
    console.warn("Error obteniendo nombre de proyecto: " + e);
  }
  return "PROYECTO GENERAL";
}

/**
 * Lógica común de cierre: Vistas, Marcado y Alertas.
 */
function finalizarEjecucionComun(ss, trabajosProcesados, idsParaMarcar, horaInicio, nombreModo) {
  const ui = SpreadsheetApp.getUi();
  const tiempoTotal = ((new Date() - horaInicio) / 1000).toFixed(1);

  if (trabajosProcesados > 0) {
    // 1. Marcar como realizados en la hoja de datos
    if (idsParaMarcar.length > 0) {
      console.log("\n--- Marcando trabajos como REALIZADOS ---");
      marcarCortesComoHechos(ss.getSheetByName(HOJA_DATOS_NOMBRE), idsParaMarcar);
    }

    // 2. Alerta Final
    let mensaje = `✅ ¡${nombreModo} Completado en ${tiempoTotal}s!\nSe procesaron ${trabajosProcesados} perfiles.`;
    if (ALERTAS.length > 0) {
      mensaje += "\n\n⚠️ RESUMEN:\n" + ALERTAS.join("\n");
    }
    ui.alert(mensaje);

  } else {
    ui.alert(`ℹ️ ${nombreModo}: No se encontraron trabajos nuevos pendientes.`);
  }
}