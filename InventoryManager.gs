/**
 * ===============================================
 * ARCHIVO: InventoryMgr.gs
 * Gestión del Inventario (Actualización, Limpieza y Stock)
 * ===============================================
 */

function limpiarInventario(hojaInv) {
  console.log("[Fase 3] Limpiando inventario de restos con cantidad 0...");
  if (hojaInv.getLastRow() < 2) return; 

  const rangoDatos = hojaInv.getRange(2, 1, hojaInv.getLastRow() - 1, 8);
  const datos = rangoDatos.getValues();

  const datosFiltrados = datos.filter(fila => {
    const tipo = fila[3];    
    const cantidad = fila[5]; 
    if (cantidad != 0 || tipo === 'Barra Nueva') return true;
    return false;
  });

  rangoDatos.clearContent();

  if (datosFiltrados.length > 0) {
    hojaInv.getRange(2, 1, datosFiltrados.length, datosFiltrados[0].length).setValues(datosFiltrados);
  }
}

function cargarEtiquetasExistentes(hojaInv) {
  const etiquetasPorPerfil = new Map();
  if (hojaInv.getLastRow() < 2) return etiquetasPorPerfil;

  const rangoDatos = hojaInv.getRange(2, 1, hojaInv.getLastRow() - 1, 7); 
  const datos = rangoDatos.getValues();

  const perfilCodeCache = new Map();
  TODOS_PERFILES.forEach(p => {
    perfilCodeCache.set(p.nombre, p.id.split('-')[1]);
  });
  
  for (const fila of datos) {
    const perfilNombre = fila[2]; 
    const etiqueta = fila[6];    
    
    if (!perfilNombre || !etiqueta) continue;

    const idPerfilBase = perfilCodeCache.get(perfilNombre);
    if (!idPerfilBase) continue;
    
    const regex = new RegExp(`^${idPerfilBase}(\\d+)$`);
    const match = etiqueta.match(regex);

    if (match) {
      const numero = parseInt(match[1], 10);
      if (!etiquetasPorPerfil.has(idPerfilBase)) {
        etiquetasPorPerfil.set(idPerfilBase, new Set());
      }
      etiquetasPorPerfil.get(idPerfilBase).add(numero);
    }
  }
  return etiquetasPorPerfil;
}

function encontrarNumeroMasBajo(numerosUsados) {
  let numeroNuevo = 1;
  while (numerosUsados.has(numeroNuevo)) {
    numeroNuevo++;
  }
  return numeroNuevo;
}

function actualizarInventario(ss, materialUsado, nuevosRestos) {
  console.log("[Fase 3] Actualizando Inventario...");
  const hojaInv = ss.getSheetByName(HOJA_INVENTARIO_NOMBRE);
  
  for (const barraUsada of materialUsado) {
    if (barraUsada.esVirtual) continue; 

    const fila = barraUsada.filaInventario;
    const celdaCantidad = hojaInv.getRange(fila, 6); 
    const cantidadActual = celdaCantidad.getValue();
    
    if (cantidadActual > 0) {
      celdaCantidad.setValue(cantidadActual - 1);
    } else {
      console.error(`Error de Inventario: Se intentó usar la fila ${fila} pero la cantidad ya era 0.`);
    }
  }
  
  limpiarInventario(hojaInv);
  const etiquetasExistentes = cargarEtiquetasExistentes(hojaInv);
  const filasNuevas = [];
  const fechaHoy = new Date();
  
  for (const resto of nuevosRestos) {
    const perfilInfo = TODOS_PERFILES.find(p => p.nombre === resto.perfil);
    if (!perfilInfo) continue;
    
    const prefijoTipo = (resto.tipo === 'Retal' ? 'R' : 'A'); 
    const idPerfilBase = perfilInfo.id.split('-')[1]; 
    const serie = "GP60";
    
    if (!etiquetasExistentes.has(idPerfilBase)) {
      etiquetasExistentes.set(idPerfilBase, new Set());
    }
    const numerosUsados = etiquetasExistentes.get(idPerfilBase);
    const numeroNuevo = encontrarNumeroMasBajo(numerosUsados);
    numerosUsados.add(numeroNuevo);

    const nuevaEtiqueta = `${idPerfilBase}${numeroNuevo}`; 
    const nuevoId = `${prefijoTipo}_${serie}-${nuevaEtiqueta}`; 

    filasNuevas.push([
      nuevoId,                   
      serie,                     
      resto.perfil,              
      resto.tipo,                
      resto.longitud.toFixed(1), 
      1,                         
      nuevaEtiqueta,             
      fechaHoy                   
    ]);
  }
  
  if (filasNuevas.length > 0) {
    hojaInv.getRange(hojaInv.getLastRow() + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
  }
}

function comprobarStockMinimo(ss, nombrePerfil) {
  const hojaInv = ss.getSheetByName(HOJA_INVENTARIO_NOMBRE);
  const claveConfig = "Min_Stock_" + nombrePerfil.replace(/ /g, "_"); 
  const stockMinimo = CONFIG[claveConfig];
  
  if (!stockMinimo) return null; 
  
  const perfilInfo = TODOS_PERFILES.find(p => p.nombre === nombrePerfil);
  if (!perfilInfo) return null;
  const idStock = perfilInfo.id; 
  
  const rangoDatos = hojaInv.getRange("A2:F" + hojaInv.getLastRow()).getValues();
  let stockRestante = 0;
  
  for (let i = 0; i < rangoDatos.length; i++) {
    if (rangoDatos[i][0] === idStock) { 
      stockRestante = parseFloat(rangoDatos[i][5]); 
      if (isNaN(stockRestante)) stockRestante = 0;
      break; 
    }
  }
  
  if (stockRestante < stockMinimo) {
    const faltan = stockMinimo - stockRestante; 
    return `🚨 ¡STOCK BAJO! Quedan ${stockRestante} barras. Mínimo: ${stockMinimo}. (Pedir ${faltan})`;
  }
  return `✅ Stock OK (${stockRestante} restantes).`;
}

function abrirPanelInventario() {
  SpreadsheetApp.getUi().alert("Función 'abrirPanelInventario' no implementada.");
}