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
      .addItem('👁️ Cambiar Vista', 'actualizarVista')
      .addItem('🎨 Verificar Colores', 'verificarEstructuraYColores') 
      .addSeparator()
      .addItem('📄 Generar PDF Taller', 'generarPDFTaller') // <--- AÑADE ESTA LÍNEA
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
  const ui = SpreadsheetApp.getUi();
  const horaInicioTotal = new Date();
  ALERTAS = []; 
  
  let resumenCompra = new Map(); 
  let todosLosIdsParaMarcar = []; 
  let todosLosRetalesGenerados = [];

  try {
    console.log("--- 🧮 INICIO MODO SOLO (PRODUCCIÓN SIN STOCK) ---");
    
    CONFIG = leerConfiguracion(ss);
    const hojaDatos = ss.getSheetByName(HOJA_DATOS_NOMBRE);
    const hojaHistorico = prepararHojaHistorico(ss);
    prepararHojaHistorico(ss);

    // --- PASO 1: ANÁLISIS Y ORDENACIÓN DE TAREAS ---
    console.log("--- 📋 Analizando complejidad de tareas... ---");
    let tareasPendientes = [];
    const longitudBarraUtil = CONFIG.Longitud_Util;

    for (const configPerfil of TODOS_PERFILES) {
      const trabajoPrevio = leerCortesPendientes(hojaDatos, configPerfil);
      const numCortes = trabajoPrevio.cortes.length;
      if (numCortes > 0) {
        // Calculamos la longitud total necesaria para estimar barras mínimas
        const longitudTotalNecesaria = trabajoPrevio.cortes.reduce((sum, c) => sum + c.medida, 0);
        tareasPendientes.push({
          config: configPerfil,
          numCortes: numCortes,
          longitudTotal: longitudTotalNecesaria
        });
      }
    }
    // Ordenar de más fácil (menos cortes) a más difícil
    tareasPendientes.sort((a, b) => a.numCortes - b.numCortes);

    if (tareasPendientes.length === 0) {
       ui.alert(`ℹ️ No hay medidas pendientes (Checks desmarcados).`);
       return;
    }

    // --- PASO 2: CÁLCULO DEL PRESUPUESTO DE TIEMPO ---
    const TIEMPO_TOTAL_DISPONIBLE_SEG = 300; // 5 minutos totales para el script
    let pesoTotal = 0;
    // Usamos el cuadrado de los cortes para dar mucho más peso a las tareas complejas
    tareasPendientes.forEach(t => { t.peso = t.numCortes * t.numCortes; pesoTotal += t.peso; });
    
    console.log(`--- ⏱️ Presupuesto Total: ${TIEMPO_TOTAL_DISPONIBLE_SEG}s (Peso total: ${pesoTotal}) ---`);
    tareasPendientes.forEach(t => {
      // Reparto proporcional con un mínimo de 15s por tarea
      t.tiempoAsignado = Math.max(15, Math.round((t.peso / pesoTotal) * TIEMPO_TOTAL_DISPONIBLE_SEG));
      console.log(`   👉 ${t.config.nombre} (${t.numCortes} cortes): ${t.tiempoAsignado}s`);
    });
    console.log("------------------------------------------");


    // --- PASO 3: BUCLE PRINCIPAL DE PROCESAMIENTO ---
    let trabajosProcesados = 0;
    
    for (const tarea of tareasPendientes) { 
      const configPerfil = tarea.config;
      const nombrePerfil = configPerfil.nombre;
      const trabajo = leerCortesPendientes(hojaDatos, configPerfil); 
      const numCortes = trabajo.cortes.length;
      
      console.log(`\n▶️ PROCESANDO: ${nombrePerfil} (${numCortes} cortes) ---`);
      todosLosIdsParaMarcar.push(...trabajo.cortes.map(c => c.id));

      // A) Calcular Stock Virtual con HOLGURA
      // Mínimo matemático si el ajuste fuera perfecto
      let barrasMinimasTeoricas = Math.ceil(tarea.longitudTotal / longitudBarraUtil);
      
      // ¡CRÍTICO! Damos el DOBLE de barras para garantizar que el motor siempre
      // tenga la opción de abrir barra nueva en lugar de generar "zona muerta".
      // El motor es eficiente y solo usará las necesarias.
      let barrasParaIntentar = barrasMinimasTeoricas * 2;
      
      console.log(`   ℹ️ Longitud total: ${tarea.longitudTotal}mm. Mínimo teórico: ${barrasMinimasTeoricas}. Stock virtual provisto: ${barrasParaIntentar} barras.`);
      
      let stockVirtual = [];
      for (let i = 1; i <= barrasParaIntentar; i++) {
        stockVirtual.push(crearBarraVirtual(nombrePerfil, i));
      }

      // B) Llamada ÚNICA al motor con stock abundante y tiempo asignado
      // Ya no hay bucle while. Se lo juega todo a una carta con recursos de sobra.
      let solucionFinal = findSolution_v31(trabajo.cortes, stockVirtual, tarea.tiempoAsignado);

      // C) Procesar Resultados
      if (solucionFinal) {
        // Contamos cuántas barras virtuales se usaron realmente
        let barrasUsadasReales = solucionFinal.filter(b => b.esVirtual).length;
        if (barrasUsadasReales > 0) {
          resumenCompra.set(nombrePerfil, barrasUsadasReales);
        }
        
        // Escribir histórico y guardar retales
        const [nuevosRetales, materialUsado] = escribirResultadosEnHistorico(hojaHistorico, solucionFinal, nombrePerfil, trabajo.cortes.length, "MODO_SOLO");
        if (nuevosRetales && nuevosRetales.length > 0) {
          todosLosRetalesGenerados.push(...nuevosRetales);
        }
      } else {
        // Si falla incluso con el doble de barras, es un problema serio.
        console.error(`🛑 [ABORTADO] No se encontró solución para ${nombrePerfil} incluso con ${barrasParaIntentar} barras y ${tarea.tiempoAsignado}s.`);
        ALERTAS.push(`❌ ${nombrePerfil}: Fallo de optimización complejo. Revisa reglas o tiempos.`);
      }
      
      trabajosProcesados++;
    } 

    // --- PASO 4: FINALIZACIÓN ---
    if (todosLosIdsParaMarcar.length > 0) {
       generarVistaTecnicaAuto();
       dividirTablaParaImpresionClonada();
       console.log("\n--- Marcando trabajos como REALIZADOS ---");
       marcarCortesComoHechos(hojaDatos, todosLosIdsParaMarcar);
    }

    console.log("--- Generando Informe Final ---");
    generarInformeMaterial(ss, resumenCompra, todosLosRetalesGenerados);
    hojaHistorico.autoResizeColumns(1, CABECERAS_HISTORICO.length);

    const tiempoTotal = ((new Date() - horaInicioTotal) / 1000).toFixed(1);
    let mensajeFinal = `🧮 Cálculo Terminado (${tiempoTotal}s)\nSe han procesado ${trabajosProcesados} perfiles.`;
    if (ALERTAS.length > 0) {
        mensajeFinal += "\n\n⚠️ ATENCIÓN:\n" + ALERTAS.join("\n");
    }
    ui.alert(mensajeFinal);

  } catch (e) {
    console.error(e);
    ui.alert("❌ ERROR CRÍTICO MODO SOLO:\n" + e.message);
  }
}

/**
 * ==============================================================================
 * MODO PRODUCCIÓN / INVENTARIO
 * - Usa el stock real de la hoja 'Inventario'.
 * - Si no hay suficiente, añade barras virtuales.
 * - Actualiza el inventario al terminar.
 * ==============================================================================
 */
function ejecutarOptimizacion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const horaInicioTotal = new Date();
  ALERTAS = []; 
  let todosLosIdsParaMarcar = []; 
  // Tiempo límite fijo por intento en modo inventario (ej. 2 minutos)
  const TIEMPO_LIMITE_POR_INTENTO_SEG = 120; 

  try {
    console.log("--- 🚀 INICIO OPTIMIZACIÓN (MODO INVENTARIO) ---");
    
    CONFIG = leerConfiguracion(ss);
    const inventarioCompleto = leerInventario(ss);
    const configSheet = ss.getSheetByName(HOJA_CONFIG_NOMBRE);
    const inventarioSheet = ss.getSheetByName(HOJA_INVENTARIO_NOMBRE);
    const hojaDatos = ss.getSheetByName(HOJA_DATOS_NOMBRE);
    
    if (!configSheet || !inventarioSheet || !hojaDatos) throw new Error("Faltan hojas esenciales.");
    
    const hojaHistorico = prepararHojaHistorico(ss);
    const huellasExistentes = cargarHuellasExistentes(hojaHistorico);

    let trabajosProcesados = 0;
    
    for (const configPerfil of TODOS_PERFILES) { 
      const nombrePerfil = configPerfil.nombre;
      const trabajo = leerCortesPendientes(hojaDatos, configPerfil);
      
      if (trabajo.cortes.length === 0) continue;
      console.log(`--- 🔄 Procesando: ${nombrePerfil} (${trabajo.cortes.length} cortes) ---`);
      
      if (huellasExistentes.has(trabajo.huella)) {
        console.warn(`[Saltado] Plan idéntico ya existe para ${nombrePerfil}.`);
        ALERTAS.push(`--- ${nombrePerfil} ---\nℹ️ Plan saltado (ya existe en histórico).`);
        continue; 
      }

      ALERTAS.push(`--- ${nombrePerfil} ---`);
      
      // Preparar stock real
      let stockDisponible = inventarioCompleto.filter(item => item.perfil === nombrePerfil);
      stockDisponible = sortStock(stockDisponible); 

      // Bucle de intentos añadiendo barras virtuales si hace falta
      let barrasNuevasVirtuales = 0;
      let solucionFinal = null;
      let stockParaIntentar = [...stockDisponible]; 

      // Límite de seguridad para barras virtuales
      const LIMITE_BARRAS_VIRTUALES = trabajo.cortes.length + 2;

      while (solucionFinal == null && barrasNuevasVirtuales <= LIMITE_BARRAS_VIRTUALES) {
        horaInicioAlgoritmo = new Date(); 
        // Pasamos el tiempo límite fijo para este modo
        let solucionIntento = findSolution_v31(trabajo.cortes, stockParaIntentar, TIEMPO_LIMITE_POR_INTENTO_SEG);
        
        if (solucionIntento) {
          solucionFinal = solucionIntento;
        } else {
          barrasNuevasVirtuales++;
          if (barrasNuevasVirtuales > LIMITE_BARRAS_VIRTUALES) {
             console.error(`🛑 [ABORTADO INVENTARIO] Límite de barras virtuales excedido para ${nombrePerfil}.`);
             ALERTAS.push(`❌ ${nombrePerfil}: No se encontró solución. Revisa el stock o las reglas.`);
             break;
          }
          stockParaIntentar.push(crearBarraVirtual(nombrePerfil, barrasNuevasVirtuales));
        }
      } 

      if (solucionFinal) {
        if (barrasNuevasVirtuales > 0) {
          ALERTAS.push(`❗️ Se necesitaron ${barrasNuevasVirtuales} BARRAS NUEVAS adicionales.`);
        }
        const [nuevosRetales, materialUsado] = escribirResultadosEnHistorico(hojaHistorico, solucionFinal, nombrePerfil, trabajo.cortes.length, trabajo.huella);
        actualizarInventario(ss, materialUsado, nuevosRetales);
        
        const alertaStock = comprobarStockMinimo(ss, nombrePerfil);
        if (alertaStock) ALERTAS.push(alertaStock);
        
        todosLosIdsParaMarcar.push(...trabajo.cortes.map(c => c.id)); 
        trabajosProcesados++;
      }
    } 
    
    if (todosLosIdsParaMarcar.length > 0) {
      marcarCortesComoHechos(hojaDatos, todosLosIdsParaMarcar);
    }

    SpreadsheetApp.flush();
    aplicarFormatosInventario(configSheet, inventarioSheet); 
    SpreadsheetApp.flush(); 

    const tiempoTotal = ((new Date() - horaInicioTotal) / 1000).toFixed(1);
    console.log(`--- 🏁 FIN OPTIMIZACIÓN en ${tiempoTotal}s ---`);
    
    if (trabajosProcesados > 0) {
        generarVistaTecnicaAuto();
        dividirTablaParaImpresionClonada();
        ui.alert(`✅ ¡Optimización Completada en ${tiempoTotal}s!\n\nRESUMEN:\n` + ALERTAS.join("\n"));

    } else {
        ui.alert(`ℹ️ No se encontraron trabajos nuevos pendientes.`);
    }

  } catch (e) {
    console.error(e);
    ui.alert("❌ ERROR MODO INVENTARIO:\n" + e.message);
  }
}