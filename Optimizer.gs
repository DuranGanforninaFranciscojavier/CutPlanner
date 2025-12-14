/**
 * ====================================================================
 * ARCHIVO: Optimizer.gs (Versión Final - Modelo "Menos es Más" Progresivo)
 * ====================================================================
 * DISEÑO DEL USUARIO:
 * 1. COSTE PROGRESIVO: Usar una barra cuesta puntos proporcionalmente a su tamaño inicial.
 * (Usar barra nueva duele más que usar un retal pequeño).
 * 2. VALOR INVERTIDO DEL RESTO: Cuanto MÁS PEQUEÑO sea el resto válido, MEJOR.
 * (Un resto de 1m vale mucho más que uno de 6m).
 * 3. VETO ZONA MUERTA: Se mantiene la penalización masiva para restos de 20cm a 1m.
 */

// --- CONFIGURACIÓN GENERAL ---
const TIEMPO_INICIO_DECAIMIENTO_SEG = 15;
const INTERVALO_SEGUNDOS_LOG = 5;
const HAZ_BASE = 2;
const HAZ_MAXIMO = 8;
const SENSIBILIDAD_DUDA_PUNTOS = 100;

// --- CONFIGURACIÓN DE PUNTUACIÓN (EL NUEVO JUEZ) ---
const PUNTAJE_MAXIMO_PERFECCION = 1200; // Resto 0
const PUNTAJE_BASURA = 1000;            // Resto < 200mm
const PUNTAJE_VETO = -20000;            // Resto 200mm - 1000mm
const PUNTAJE_RETAL_PEQUENO = 800;      // Resto = 1000mm (El mejor retal)
const PUNTAJE_RETAL_GRANDE = 100;       // Resto = 6500mm (El peor retal)

const COSTE_BASE_BARRA_NUEVA = -1000;   // Coste por usar 6500mm de material


// ==========================================================================
// SECCIÓN 1: UTILIDADES Y SISTEMA DE PUNTUACIÓN
// ==========================================================================

function sortStock(stock) {
  return stock.sort((a, b) => {
    // Best Fit: Preferir siempre las barras más pequeñas donde quepa la pieza
    return a.longitud - b.longitud;
  });
}

function crearBarraVirtual(nombrePerfil, numeroVirtual) {
  return {
    filaInventario: null, etiquetaMaterial: `VIRTUAL-${numeroVirtual}`,
    perfil: nombrePerfil, longitud: CONFIG.Longitud_Barra_Nueva,
    tipo: "Barra Nueva", esVirtual: true, espacioRestante: CONFIG.Longitud_Util, cortes: []
  };
}

/**
 * EL JUEZ FINAL: Implementa la lógica de Coste Progresivo y Valor Invertido.
 */
function calcularPuntuacion(barra, restoResultante) {
  // --- PARTE 1: COSTE PROGRESIVO POR USO DE MATERIAL ---
  // Cuanto más larga sea la barra original que estamos "gastando", más duele.
  // Fórmula: -1000 * (Longitud / 6500)
  const factorUso = barra.longitud / CONFIG.Longitud_Barra_Nueva;
  const pCosteMaterial = COSTE_BASE_BARRA_NUEVA * factorUso;


  // --- PARTE 2: VALOR INVERTIDO DEL RESTO GENERADO ---
  let pCalidadResto = 0;
  const R = restoResultante;
  const L_BASURA = CONFIG.Limite_Desechable; // 200mm
  const L_RETAL_MIN = CONFIG.Limite_Retal;   // 1000mm
  const L_RETAL_MAX = CONFIG.Longitud_Barra_Nueva; // 6500mm

  if (R === 0) {
      // PERFECCIÓN ABSOLUTA
      pCalidadResto = PUNTAJE_MAXIMO_PERFECCION; // +1200
  } 
  else if (R < L_BASURA) {
      // BASURA ÚTIL (Casi perfecto). De +1200 a +1000.
      const factor = R / L_BASURA; // 0 -> 1
      pCalidadResto = PUNTAJE_MAXIMO_PERFECCION - ((PUNTAJE_MAXIMO_PERFECCION - PUNTAJE_BASURA) * factor);
  } 
  else if (R < L_RETAL_MIN) {
      // ZONA MUERTA (VETO)
      pCalidadResto = PUNTAJE_VETO; // -20000
  } 
  else {
      // TRAMO RETAL ÚTIL (Invertido: cuanto más pequeño, mejor)
      // Cae linealmente desde +800 (en 1m) hasta +100 (en 6.5m)
      const rangoRetal = L_RETAL_MAX - L_RETAL_MIN;
      const posicionEnRango = R - L_RETAL_MIN;
      const factorDescuento = posicionEnRango / rangoRetal; // 0 (en 1m) -> 1 (en 6.5m)
      
      const caidaTotalPuntos = PUNTAJE_RETAL_PEQUENO - PUNTAJE_RETAL_GRANDE; // 800 - 100 = 700
      pCalidadResto = PUNTAJE_RETAL_PEQUENO - (caidaTotalPuntos * factorDescuento);
  }

  // Puntuación Final = Coste (negativo) + Calidad Resto (positivo/negativo)
  return pCosteMaterial + pCalidadResto;
}

/**
 * Calcula el TECHO y el SUELO dinámicos para el rango de paciencia.
 */
function calcularRangoPaciencia(stock) {
  // TECHO: La mejor jugada posible.
  // Usar el retal más pequeño disponible (menor coste) y lograr Resto 0 (+1200).
  const retalMasPequeno = stock.reduce((min, b) => b.longitud < min.longitud ? b : min, stock[0]);
  const mejorCoste = COSTE_BASE_BARRA_NUEVA * (retalMasPequeno.longitud / CONFIG.Longitud_Barra_Nueva);
  const techo = mejorCoste + PUNTAJE_MAXIMO_PERFECCION;

  // SUELO: La peor jugada aceptable.
  // Abrir barra nueva (-1000) y dejar el peor retal válido (+100).
  const suelo = COSTE_BASE_BARRA_NUEVA + PUNTAJE_RETAL_GRANDE; // -1000 + 100 = -900

  return { techo, suelo, rango: techo - suelo };
}


// ==========================================================================
// SECCIÓN 2: EL MOTOR CEREBRAL
// ==========================================================================

function findSolution_v31(allCuts, stockDisponible, tiempoLimiteSegundos) {
  allCuts.sort((a, b) => b.medida - a.medida);
  const numCortes = allCuts.length;
  const numBarras = stockDisponible.length;
  
  let barrasTrabajo = stockDisponible.map(s => ({
    ...s, espacioRestante: s.esVirtual ? CONFIG.Longitud_Util : s.longitud, cortes: []
  }));

  // Calibración del Rango de Paciencia (Basado en tu nuevo modelo)
  const { techo, suelo, rango } = calcularRangoPaciencia(stockDisponible);

  const horaInicioGlobal = new Date();
  let nodosExploradosTotal = 0;
  let ultimoLogTiempoSeg = -INTERVALO_SEGUNDOS_LOG;

  console.log(`🧠 [MOTOR] Iniciando: ${numCortes} cortes en ${numBarras} barras.`);
  console.log(`🎯 [CALIBRACIÓN] Techo: ${techo.toFixed(0)} | Suelo: ${suelo.toFixed(0)} | Rango: ${rango.toFixed(0)}.`);

// --- FUNCIÓN RECURSIVA CENTRAL (CORREGIDA Y BLINDADA) ---
  function explorar(cutIndex) {
    nodosExploradosTotal++;
    const tiempoPasado = (new Date() - horaInicioGlobal) / 1000;

    // A. SEGURIDAD GOOGLE
    if (tiempoPasado > (tiempoLimiteSegundos + 5)) throw new Error("TIMEOUT_GOOGLE_SAFETY");

    // B. CASO BASE (Todos los cortes colocados)
    if (cutIndex === allCuts.length) return true;

    // C. CÁLCULO DE LA PACIENCIA ACTUAL (Interpolación lineal Techo -> Suelo)
    let umbralActual = techo;
    if (tiempoPasado > TIEMPO_INICIO_DECAIMIENTO_SEG) {
        const tiempoEnDecaimiento = Math.min(tiempoPasado, tiempoLimiteSegundos) - TIEMPO_INICIO_DECAIMIENTO_SEG;
        const duracionTotal = tiempoLimiteSegundos - TIEMPO_INICIO_DECAIMIENTO_SEG;
        const factor = duracionTotal > 0 ? tiempoEnDecaimiento / duracionTotal : 1;
        umbralActual = techo - (rango * factor);
    }

    // LOGGING (Solo cada X segundos)
    if (tiempoPasado - ultimoLogTiempoSeg >= INTERVALO_SEGUNDOS_LOG) {
       console.log(`⏱️ [ESTADO] T=${tiempoPasado.toFixed(1)}s/${tiempoLimiteSegundos}s | Nodos: ${nodosExploradosTotal} | Paciencia: ${umbralActual.toFixed(0)} pts`);
       ultimoLogTiempoSeg = tiempoPasado;
    }

    // D. GENERACIÓN Y EVALUACIÓN DE CANDIDATOS (EL FIX ESTÁ AQUÍ)
    const corteActual = allCuts[cutIndex];
    // 1. Calculamos la longitud REAL que ocupa el corte (Medida + Desperdicio)
    const longitudOcupada = corteActual.medida + CONFIG.Desperdicio_por_Corte;

    let candidatos = [];
    for (let i = 0; i < barrasTrabajo.length; i++) {
      
      // 2. Calculamos el resto potencial usando la longitud OCUPADA
      const restoPotencial = barrasTrabajo[i].espacioRestante - longitudOcupada;

      // --- BLINDAJE FÍSICO ---
      // 3. Si el resto potencial es negativo (más allá de una mínima tolerancia por decimales),
      // la pieza NO CABE físicamente. Saltamos esta barra inmediatamente.
      if (restoPotencial < -0.001) {
        continue; 
      }
      
      // Si llegamos aquí, es que la pieza cabe y el resto es >= 0.
      // Procedemos a puntuar la opción.
      const score = calcularPuntuacion(barrasTrabajo[i], restoPotencial);
      // Guardamos 'restoPotencial' como el 'resto' real que quedaría
      candidatos.push({ indexBarra: i, score: score, resto: restoPotencial });
    }    

    // Si no cabe en ninguna barra, este camino ha fallado.
    if (candidatos.length === 0) return false;

    // Ordenar candidatos de mejor a peor puntuación
    candidatos.sort((a, b) => b.score - a.score);
    
    // VETO ZONA MUERTA (Seguridad adicional rápida)
    // Si la mejor opción sigue siendo horrible, cortamos por lo sano.
    if (candidatos[0].score < PUNTAJE_VETO + 5000) return false;

    // E. SELECCIÓN ADAPTATIVA DEL HAZ (Beam Search)
    let hazAExplorar = [];
    // Regla 1: Aceptación Temprana (Greedy Dinámico)
    if (candidatos[0].score >= umbralActual) {
       hazAExplorar = [candidatos[0]];
    } 
    // Regla 2: Incertidumbre (Haz Respirable)
    else {
       let anchoHaz = HAZ_BASE;
       for (let k = 1; k < Math.min(candidatos.length, HAZ_MAXIMO); k++) {
         // Si la diferencia de puntos es pequeña, ampliamos el haz
         if (candidatos[0].score - candidatos[k].score < SENSIBILIDAD_DUDA_PUNTOS) {
           anchoHaz = k + 1;
         } else {
           break;
         }
       }
       hazAExplorar = candidatos.slice(0, anchoHaz);
    }

    // F. BUCLE DE EXPLORACIÓN RECURSIVA
    for (const mov of hazAExplorar) {
      const barra = barrasTrabajo[mov.indexBarra];
      const espacioOriginal = barra.espacioRestante;
      
      // Aplicar movimiento
      barra.espacioRestante = mov.resto;
      barra.cortes.push(corteActual);

      // RECURSIÓN: Intentar colocar el siguiente corte
      if (explorar(cutIndex + 1)) return true;

      // BACKTRACKING: Deshacer movimiento si el camino falló
      barra.cortes.pop();
      barra.espacioRestante = espacioOriginal;
    }

    return false; // Ningún camino del haz funcionó
  }

  try {
    if (explorar(0)) {
      let puntuacionTotalSolucion = 0;
      barrasTrabajo.forEach(b => {
        if (b.cortes.length > 0) puntuacionTotalSolucion += calcularPuntuacion(b, b.espacioRestante);
      });
      const tiempoTotal = (new Date() - horaInicioGlobal) / 1000;
      console.log(`🏁 [ÉXITO MOTOR] Solución en ${tiempoTotal.toFixed(1)}s. Nodos: ${nodosExploradosTotal}. Score: ${puntuacionTotalSolucion.toFixed(0)}.`);
      return barrasTrabajo.filter(b => b.cortes.length > 0);
    } else {
      console.warn(`⛔ [FALLO MOTOR] Sin solución tras ${nodosExploradosTotal} nodos.`);
      return null;
    }
  } catch (e) {
    if (e.message === "TIMEOUT_GOOGLE_SAFETY") {
        console.error(`⏱️ [TIMEOUT] Tiempo agotado. Nodos: ${nodosExploradosTotal}.`);
        return null; 
    } else { throw e; }
  }
}