/**
 * ==============================================================================
 * ARCHIVO: UIManager.gs
 * GESTIÓN DE INTERFAZ DE USUARIO Y EVENTOS
 * ==============================================================================
 * Contiene:
 * 1. El disparador principal onEdit.
 * 2. Funciones para gestionar las vistas de la hoja de entrada ("Corredera_GP-60").
 * 3. Funciones para gestionar las vistas de la hoja de resultados ("Plan_Cortes").
 * 4. Lógica de cálculo automático de medidas de ventana.
 */

const NOMBRE_HOJA_TRABAJO = "Corredera_GP-60"; // Nombre de tu pestaña de entrada

// ==============================================================================
// SECCIÓN 1: DISPARADOR AUTOMÁTICO AL EDITAR (onEdit)
// ==============================================================================
/**
 * Se ejecuta automáticamente cada vez que se edita una celda en la hoja de cálculo.
 * Actúa como un "router", dirigiendo la acción según la hoja donde ocurrió la edición.
 */
function onEdit(e) {
  // Guardas de seguridad
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

console.log("Edit")
  // --- CASO 1: EDICIÓN EN LA HOJA DE TRABAJO (ENTRADA DE DATOS) ---
  if (sheetName === NOMBRE_HOJA_TRABAJO) {
    // A) Cambio del desplegable de vista en A1 (Fila 1, Col 1)
    if (row === 1 && col === 1) {
      actualizarVistaTrabajo(); // Función renombrada para claridad
      return; // Salimos para no ejecutar cálculos innecesarios
    }

    // B) Edición de datos que requiere recálculo (Filas > 2)
    if (row > 2) {
      // Columnas clave: B(2, L), C(3, H), E(5, Tipo), F(6, T)
      if (col === 2 || col === 3 || col === 5 || col === 6) {
        calcularFila(sheet, row);
      }
    }
  }
  
  // --- CASO 2: EDICIÓN EN LA HOJA DE CONFIGURACIÓN ---
  else if (sheetName === HOJA_CONFIG_NOMBRE) {
          console.log("Cambios de Configuración en:"+ row + ":"+col)

    // A) Cambio del desplegable de vista del plan de corte en D3 (Fila 3, Col 4)

    if (row === 4 && col === 5) {
      console.log("Cambios de plan de corte")
      aplicarVistaPlanCortes(e.value);
    }
  }
}


// ==============================================================================
// SECCIÓN 2: VISTAS DE LA HOJA DE TRABAJO ("Corredera_GP-60")
// ==============================================================================

/**
 * FUNCIÓN AUXILIAR PRIVADA PARA HOJA DE TRABAJO
 * Devuelve la configuración de columnas para las vistas.
 */
function _getConfigVistaTrabajo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_TRABAJO);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Error: No se encuentra la hoja '${NOMBRE_HOJA_TRABAJO}'.`);
    return null;
  }
  return {
    sheet: sheet,
    colInicio: 4,   // Columna D
    totalCols: 11   // D hasta N
  };
}

// --- MÉTODOS INDEPENDIENTES (Se pueden asignar a botones) ---

function vistaTodoTrabajo() {
  const config = _getConfigVistaTrabajo();
  if (!config) return;
  config.sheet.showColumns(config.colInicio, config.totalCols);
  SpreadsheetApp.getActive().toast("Vista Completa activada", "👁️ Hoja Trabajo");
}

function vistaBaseTrabajo() {
  const config = _getConfigVistaTrabajo();
  if (!config) return;
  config.sheet.hideColumns(config.colInicio, config.totalCols);
  SpreadsheetApp.getActive().toast("Vista Base activada", "👁️ Hoja Trabajo");
}

function vistaTecnicoTrabajo() {
  const config = _getConfigVistaTrabajo();
  if (!config) return;
  // Mostrar D, E, F (3 columnas)
  config.sheet.showColumns(config.colInicio, 3);
  // Ocultar G hasta N (8 columnas restantes)
  config.sheet.hideColumns(config.colInicio + 3, config.totalCols - 3);
  SpreadsheetApp.getActive().toast("Vista Técnica (D-F visibles)", "👁️ Hoja Trabajo");
}

/**
 * Función principal que llama onEdit al cambiar el desplegable en A1.
 * Renombrada de 'actualizarVista' a 'actualizarVistaTrabajo' para evitar confusión.
 */
function actualizarVistaTrabajo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_TRABAJO);
  const modo = sheet.getRange("A1").getValue();

  if (modo === "Todo") {
    vistaTodoTrabajo();
  } else if (modo === "Base") {
    vistaBaseTrabajo();
  } else if (modo === "Tecnico") {
    vistaTecnicoTrabajo();
  }
}


// ==============================================================================
// SECCIÓN 3: VISTAS DE LA HOJA DE PLAN DE CORTES (Histórico)
// ==============================================================================

/**
 * FUNCIÓN AUXILIAR PRIVADA PARA PLAN DE CORTES
 * Verifica la hoja y resetea la vista mostrando todo antes de aplicar filtros.
 */
function _prepararHojaPlanCortes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaHistorico = ss.getSheetByName(HOJA_HISTORICO_NOMBRE);

  if (!hojaHistorico) {
    // Si no existe la hoja, no mostramos alerta molesta, solo salimos.
    // Puede ser normal si aún no se ha ejecutado ninguna optimización.
    return null;
  }

  // 1. Mostrar TODO primero para asegurar un estado limpio
  const maxCols = hojaHistorico.getLastColumn();
  if (maxCols > 0) {
      hojaHistorico.showColumns(1, maxCols);
  }
  
  // 2. Ajustar ancho de la columna de Cortes (G, índice 7) para mejor lectura
  hojaHistorico.setColumnWidth(7, 250); 

  return hojaHistorico;
}

// --- MÉTODOS INDEPENDIENTES (Se pueden asignar a botones) ---

function vistaTodoPlanCorte() {
  const hoja = _prepararHojaPlanCortes();
  if (!hoja) return;
  // No hay que ocultar nada, _prepararHojaVista ya lo mostró todo.
  SpreadsheetApp.getActive().toast("Mostrando todas las columnas.", "👁️ Plan Cortes: TODO");
}

function vistaTallerSoloPlanCorte() {
    console.log("Vista Simple Plan corte");

  const hoja = _prepararHojaPlanCortes();
  if (!hoja) return;
  // A(1): ID, D(4): Etiqueta, F(6): Long. Inicial
// Ocultar A(1) y B(2) -> Desde la 1, 2 columnas
      hoja.hideColumns(1, 2);
      
      // Ocultar D(4) -> Desde la 4, 1 columna
      hoja.hideColumns(4, 2);

      // Ocultar F(6) -> Desde la 6, 1 columna
      hoja.hideColumns(6, 1);

      // Ocultar L(12) (Huella) -> Desde la 12, 1 columna
      hoja.hideColumns(12, 1);

  // Bloque final: I(9) a M(13) - 5 columnas
  hoja.hideColumns(9, 5); 

  SpreadsheetApp.getActive().toast("Vista simplificada para taller aplicada.", "👁️ Plan Cortes: TALLER");
}

function vistaInventario() {
  const hoja = _prepararHojaPlanCortes();
  if (!hoja) return;

  // A(1): ID
  hoja.hideColumn(1);
  
  // Bloque final: I(9) a M(13) - 5 columnas
  hoja.hideColumns(9, 5);

  SpreadsheetApp.getActive().toast("Vista de inventario y cortes aplicada.", "👁️ Plan Cortes: INVENTARIO");
}

/**
 * Función auxiliar que aplica la vista seleccionada desde el desplegable de Configuración.
 * Se llama desde onEdit.
 */
function aplicarVistaPlanCortes(vistaSeleccionada) {
  // Esta función ahora es muy simple, solo llama a los métodos modulares.
  switch (vistaSeleccionada) {
    case 'Completa':
      vistaTodoPlanCorte();
      break;
    case 'Base':
      vistaTallerSoloPlanCorte();
      break;
    default:
      // Si el valor no es uno de los esperados, no hacemos nada.
      return;
  }
  hojaHistorico.autoResizeColumns(1, CABECERAS_HISTORICO.length);
}


// ==============================================================================
// SECCIÓN 4: CÁLCULOS DE VENTANA (Lógica de Negocio)
// ==============================================================================

/**
 * Recorre toda la hoja y fuerza el recalculo de todas las filas válidas.
 */
function actualizarTodo() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_TRABAJO);
  var ultimaFila = obtenerUltimaFilaValida(sheet);

  if (ultimaFila === 0) {
      SpreadsheetApp.getActive().toast("No se encontraron filas con datos para actualizar.");
      return;
  }

  // Bucle desde la fila 3 hasta la última válida
  for (var i = 3; i <= ultimaFila; i++) {
    calcularFila(sheet, i);
  }
  
  SpreadsheetApp.getActive().toast(`Se han actualizado ${ultimaFila - 2} filas.`, "🔄 Recálculo Completo");
}

/**
 * Encuentra la última fila que contiene datos de medidas (Ancho y Alto).
 */
function obtenerUltimaFilaValida(sheet) {
  var maxFila = sheet.getLastRow();
  if (maxFila < 3) return 0;
  
  // Leemos solo las columnas B(2) y C(3) desde la fila 3 hasta el final
  var valores = sheet.getRange(3, 2, maxFila - 2, 2).getValues();
  
  // Recorremos el array de atrás hacia adelante para encontrar el último dato
  for (var i = valores.length - 1; i >= 0; i--) {
    var ancho = valores[i][0];
    var alto = valores[i][1];
    
    // Si encontramos datos válidos, retornamos la posición real en la hoja
    if (ancho > 0 && alto > 0) {
      return i + 3; // +3 porque el array empieza en la fila 3 (índice 0 del array)
    }
  }
  return 0; // No se encontraron datos
}

/**
 * Lógica principal de cálculo de una fila.
 */
function calcularFila(sheet, fila) {
  // Leer datos clave: B(Ancho), C(Alto), E(Tipo), F(T)
  // Leemos un rango que cubra todo lo que necesitamos
  const rangoDatos = sheet.getRange(fila, 2, 1, 15); // Desde B(2) hasta F(6)
  const valores = rangoDatos.getValues()[0];

  const L = valores[13]; // Columna B (Índice 0)
  const H = valores[14]; // Columna C (Índice 1)
  console.log("L "+L + " H "+ H)
  // Columna D (Índice 2) - Saltada
  const tipo = valores[3]; // Columna E (Índice 3)
  const T = valores[4]; // Columna F (Índice 4) - Necesario para Compacto

  // Validar que haya datos suficientes para calcular.
  // El tipo es obligatorio. L y H también. T solo si es compacto, pero lo validamos en su fórmula.
  if (!L || !H || !tipo) return;

  // Definir rangos de salida
  // CORREDERA 2 HOJAS y COMPACTO (Q-V) -> Col 17, 6 celdas de ancho
  var rangoSalida = sheet.getRange(fila, 17, 1, 6);
  var resultados = [];

  if (tipo === "2 Hojas") {
    resultados = formulas2Hojas(L, H);
    rangoSalida.setValues([resultados]);
    
  } else if (tipo === "Compacto") {
    // Para compacto necesitamos que T tenga valor
    if (T) {
        resultados = formulasCompacto(L, H, T);
        rangoSalida.setValues([resultados]);
    } else {
        // Si falta T, quizás deberíamos limpiar el rango o avisar
        // rangoSalida.clearContent(); 
    }
  } else {
    // Si es otro tipo o se borra, limpiamos el rango de resultados
    rangoSalida.clearContent();
  }
}

// --- FÓRMULAS ---
function formulas2Hojas(L, H) {
  return [
    L - 37 ,    // Marco superior
    L - 37,    // Marco Inferior
    H,             // Marco Lateral
    H - 55,        // Hoja Lateral
    H - 55,        // Hoja Central
    L / 2 - 14.5   // Hoja Ruedas
  ];
}

function formulasCompacto(L, H, T) {
  console.log(L+ " "+ H+ " "+ T)
  return [
    L - 90,    // Marco superior
    L - 90,    // Marco Inferior
    H - T,         // Marco Lateral
    H - T - 55,    // Hoja Lateral
    H - T - 55,       // Hoja Central
    L / 2 - 40     // Hoja Ruedas
  ];
}