/**
 * ===============================================
 * ARCHIVO: Config.gs
 * Configuración Global, Catálogo de Perfiles y Mapeo Dinámico
 * ===============================================
 */

// --- ESTRUCTURA DE FILAS ---
const FILA_CABECERA = 1;
const FILA_INICIO_DATOS = 2;

// --- NOMBRES DE HOJAS ---
const HOJA_DATOS_NOMBRE = 'Corredera_GP-60'; 
const HOJA_INVENTARIO_NOMBRE = 'Inventario';
const HOJA_CONFIG_NOMBRE = 'Configuración';
const HOJA_HISTORICO_NOMBRE = 'Plan_Cortes';
const HOJA_RESUMEN_NOMBRE = 'Material_Necesario';
const HOJA_MATERIAL_NECESARIO = 'Material_Necesario';

/**
 * Función auxiliar para obtener un mapa dinámico de columnas según los encabezados de la fila 1.
 * Permite buscar columnas por su nombre exacto o normalizado (sin espacios, sin guiones, minúsculas).
 */
function getHeaderMap(sheet, filaCabecera = FILA_CABECERA) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return { getCol: () => null };

  const headers = sheet.getRange(filaCabecera, 1, 1, lastCol).getValues()[0];
  const map = {};

  for (let i = 0; i < headers.length; i++) {
    const raw = String(headers[i] || '').trim();
    if (raw) {
      map[raw] = i + 1;
      const norm = raw.toLowerCase().replace(/[\s_-]+/g, '');
      map[norm] = i + 1;
    }
  }

  map.getCol = function(name) {
    if (!name) return null;
    if (this[name] !== undefined) return this[name];
    const norm = String(name).toLowerCase().replace(/[\s_-]+/g, '');
    return this[norm] !== undefined ? this[norm] : null;
  };

  return map;
}

// --- DEFINICIÓN DE PERFILES (Nomenclatura con _) ---
const TODOS_PERFILES = [
  // Marcos
  { 
    id: "BN_GP60-MS", 
    nombre: "Marco_Superior", 
    columna: "Marco_Superior", 
    multiplicador: 1, 
    color: "#94DAFF", 
    nombreBase: "Marco_Superior" 
  }, 
  { 
    id: "BN_GP60-MI", 
    nombre: "Marco_Inferior", 
    columna: "Marco_Inferior", 
    multiplicador: 1, 
    color: "#FCE5CD", 
    nombreBase: "Marco_Inferior" 
  }, 
  
  // Marco Lateral (Para tipo '2 Hojas')
  { 
    id: "BN_GP60-ML", 
    nombre: "Marco_Lateral", 
    columna: "Marco_Lateral", 
    multiplicador: 2, 
    color: "#FFC494", 
    nombreBase: "Marco_Lateral",
    tipoRequerido: ["2 Hojas", "2_Hojas"] 
  },

  // Marcos Laterales Compacto (Izquierda y Derecha)
  { 
    id: "BN_GP60-MLG-IZQ", 
    nombre: "Marco_Lat_Compacto_IZQ", 
    columna: "Marco_Lateral", 
    multiplicador: 1, 
    color: "#F9CB9C", 
    nombreBase: "Marco_Lateral",
    tipoRequerido: ["Compacto"] 
  },
  { 
    id: "BN_GP60-MLG-DER", 
    nombre: "Marco_Lat_Compacto_DER", 
    columna: "Marco_Lateral", 
    multiplicador: 1, 
    color: "#E6B8AF", 
    nombreBase: "Marco_Lateral",
    tipoRequerido: ["Compacto"] 
  },

  // Perfiles Hoja H10
  { 
    id: "BN_GP60-HL10", 
    nombre: "Hoja_Lateral_H10",  
    columna: "Hoja_Lateral", 
    multiplicador: 2, 
    color: "#F299BF", 
    nombreBase: "Hoja_Lateral_H10" 
  },
  { 
    id: "BN_GP60-HC10", 
    nombre: "Hoja_Central_H10",  
    columna: "Hoja_Central", 
    multiplicador: 2, 
    color: "#BEFACA", 
    nombreBase: "Hoja_Central_H10" 
  },
  { 
    id: "BN_GP60-HR10", 
    nombre: "Hoja_Ruedas_H10",   
    columna: "Hoja_Ruedas", 
    multiplicador: 4, 
    color: "#D3B8E0", 
    nombreBase: "Hoja_Ruedas_H10" 
  },

  // Perfiles Hoja H18
  { 
    id: "BN_GP60-HL18", 
    nombre: "Hoja_Lateral_H18",  
    columna: "Hoja_Lateral", 
    multiplicador: 2, 
    color: "#BF7897", 
    nombreBase: "Hoja_Lateral_H18" 
  },
  { 
    id: "BN_GP60-HC18", 
    nombre: "Hoja_Central_H18",  
    columna: "Hoja_Central", 
    multiplicador: 2, 
    color: "#95C79F", 
    nombreBase: "Hoja_Central_H18" 
  },
  { 
    id: "BN_GP60-HR18", 
    nombre: "Hoja_Ruedas_H18",   
    columna: "Hoja_Ruedas", 
    multiplicador: 4, 
    color: "#A48BB0", 
    nombreBase: "Hoja_Ruedas_H18" 
  }
];

// --- CABECERAS COLA DE CORTES (AppSheet Ready) ---
const CABECERAS_HISTORICO = [
  "ID_Corte",
  "ID_Barra",
  "Proyecto",
  "Ventana",
  "Perfil",
  "Medida_Corte",
  "Posicion_Corte",
  "Tipo_Material",
  "Etiqueta_Material",
  "Long_Barra",
  "Resto_Barra",
  "Tipo_Resto",
  "Estado",
  "Fecha_Hora",
  "Huella"
];

// --- ÍNDICES DE COLUMNAS HISTÓRICO ---
const COL_ID_CORTE_IDX = 1;
const COL_ID_BARRA_IDX = 2;
const COL_PROYECTO_IDX = 3;
const COL_VENTANA_IDX = 4;
const COL_PERFIL_IDX = 5;
const COL_MEDIDA_IDX = 6;
const COL_POSICION_IDX = 7;
const COL_TIPO_MAT_IDX = 8;
const COL_ETIQUETA_MAT_IDX = 9;
const COL_LONG_BARRA_IDX = 10;
const COL_RESTO_BARRA_IDX = 11;
const COL_TIPO_RESTO_IDX = 12;
const COL_ESTADO_IDX = 13;
const COL_FECHA_IDX = 14;
const COL_HUELLA_IDX = 15;

// --- COLORES ---
const COLOR_FONDO_BASE = '#f3f3f3'; 
const COLOR_FONDO_CABECERA = '#666666';
const COLOR_TEXTO_CABECERA = '#ffffff';
const COLOR_TEXTO_GRIS = '#999999'; 
const COLOR_RESTO_DESECHABLE = '#f4cccc'; 
const COLOR_RESTO_RETAL = '#d9ead3';
const COLOR_RESTO_ALMACENABLE = '#cfe2f3';

const COLORES_TIPO_HISTORICO = {
  'Barra Nueva': '#DAF5F5',   
  'Retal': '#C9EAD0',       
  'Almacenable': '#C9DAF8', 
  'Pedir Barra': '#F5A293'   
};

// --- CONFIGURACIÓN DEL MOTOR ---
const LIMITE_TIEMPO_SEGUNDOS = 300; 
let horaInicioAlgoritmo;

// --- COLORES VISUALES (Inventario) ---
const COLORES_PERFIL = {
  'Marco_Superior': '#94DAFF',   
  'Marco_Inferior': '#FCE5CD',   
  'Marco_Lateral': '#FFC494',    
  'Marco_Lat_Compacto_IZQ': '#F9CB9C',
  'Marco_Lat_Compacto_DER': '#E6B8AF',
  'Hoja_Lateral_H10': '#F299BF', 
  'Hoja_Lateral_H18': '#BF7897', 
  'Hoja_Central_H10': '#BEFACA', 
  'Hoja_Central_H18': '#95C79F', 
  'Hoja_Ruedas_H10': '#D3B8E0',  
  'Hoja_Ruedas_H18': '#A48BB0',
  // Compatibilidad con nombres con espacio
  'Marco Superior': '#94DAFF',   
  'Marco Inferior': '#FCE5CD',   
  'Marco Lateral': '#FFC494',    
  'Hoja Lateral H10': '#F299BF', 
  'Hoja Lateral H18': '#BF7897', 
  'Hoja Central H10': '#BEFACA', 
  'Hoja Central H18': '#95C79F', 
  'Hoja Ruedas H10': '#D3B8E0',  
  'Hoja Ruedas H18': '#A48BB0'
};

const COLORES_TIPO = {
  'Barra Nueva': '#EEEEEE',   
  'Retal': '#C9EAD0',       
  'Almacenable': '#C9DAF8', 
  'Pedir Barra': '#FFF2CC'   
};

const COLORES_CANTIDAD = {
  'PELIGRO': '#FF0000',     
  'BAJO': '#FFA500',       
  'ADVERTENCIA': '#FFFF00',    
  'EXCESO': '#9900FF'       
};

// Variable global para alertas y configuración en memoria
let CONFIG = {};
let ALERTAS = [];