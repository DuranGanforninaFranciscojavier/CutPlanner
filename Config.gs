
// --- COLUMNAS DE ENTRADA ---
const COLUMNA_ID_LETRA = "H";          // ID no suele ser la B (Revisa esto abajo*)
const COLUMNA_HUECO_LETRA = "D";       // Aquí pone "H10", "H18"...
const COLUMNA_TIPO_LETRA = "E";        // Aquí pone "2 Hojas", "Compacto"...
const COLUMNA_CALCULADO_LETRA = "I";   // El Checkbox maestro: TRUE/FALSE
const COLUMNA_HECHO_LETRA = "I";
const FILA_INICIO_DATOS = 3;
const HOJA_RESUMEN_NOMBRE = 'Material_Necesario';

// --- DEFINICIÓN DE PERFILES (ACTUALIZADA) ---
// Se ha añadido la propiedad 'tipoRequerido' para filtrar qué perfil se activa.
const TODOS_PERFILES = [
  // Marcos
  { id: "BN_GP60-MS", nombre: "Marco Superior",    columna: "Q", multiplicador: 1, color: "#94DAFF", nombreBase: "Marco Superior" }, 
  { id: "BN_GP60-MI", nombre: "Marco Inferior",    columna: "R", multiplicador: 1, color: "#FCE5CD", nombreBase: "Marco Inferior" }, 
  
  // --- MARCO LATERAL (PARA TIPO '2 Hojas') ---
  { 
    id: "BN_GP60-ML", 
    nombre: "Marco Lateral",        // Nombre del material normal
    columna: "S", 
    multiplicador: 2, 
    color: "#FFC494", 
    nombreBase: "Marco Lateral",
    // NUEVO: Solo se activa si la columna 'Tipo' (E) contiene "2 Hojas"
    tipoRequerido: ["2 Hojas"] 
  },

  // --- NUEVO PERFIL: MARCO LATERAL GUÍA (PARA TIPO 'Compacto') ---
  // Lee la misma columna 'S', pero es un material distinto.
// 1. MARCO LATERAL IZQUIERDO (Compacto)
  { 
    id: "BN_GP60-MLG-IZQ",        // ID Único para el stock
    nombre: "Marco Lat. Compacto IZQ", // Nombre para mostrar
    columna: "S",                 // Lee la misma medida de altura
    multiplicador: 1,             // SOLO 1 UNIDAD
    color: "#F9CB9C",             // Color naranja claro
    nombreBase: "Marco Lat. Compacto IZQ",
    tipoRequerido: ["Compacto"]   // Solo si es Compacto
  },

  // 2. MARCO LATERAL DERECHO (Compacto)
  { 
    id: "BN_GP60-MLG-DER",        // ID Único diferente
    nombre: "Marco Lat. Compacto DER", 
    columna: "S",                 // Lee la misma medida de altura
    multiplicador: 1,             // SOLO 1 UNIDAD
    color: "#E6B8AF",             // Un color ligeramente diferente para distinguir
    nombreBase: "Marco Lat. Compacto DER",
    tipoRequerido: ["Compacto"]   // Solo si es Compacto
  },

  // H10
  { id: "BN_GP60-HL10", nombre: "Hoja Lateral H10",  columna: "T", multiplicador: 2, color: "#F299BF", nombreBase: "Hoja Lateral H10" },
  { id: "BN_GP60-HC10", nombre: "Hoja Central H10",  columna: "U", multiplicador: 2, color: "#BEFACA", nombreBase: "Hoja Central H10" },
  { id: "BN_GP60-HR10", nombre: "Hoja Ruedas H10",   columna: "V", multiplicador: 4, color: "#D3B8E0", nombreBase: "Hoja Ruedas H10" },

  // --- H18 ---
  { id: "BN_GP60-HL18", nombre: "Hoja Lateral H18",  columna: "T", multiplicador: 2, color: "#BF7897", nombreBase: "Hoja Lateral H18" },
  { id: "BN_GP60-HC18", nombre: "Hoja Central H18",  columna: "U", multiplicador: 2, color: "#95C79F", nombreBase: "Hoja Central H18" },
  { id: "BN_GP60-HR18", nombre: "Hoja Ruedas H18",   columna: "V", multiplicador: 4, color: "#A48BB0", nombreBase: "Hoja Ruedas H18" }
];

// --- CABECERAS HISTÓRICO ---
const CABECERAS_HISTORICO = [
  "ID", "Serie", "Perfil", "Etiqueta Mat", 
  "Tipo Material", "Long. Inicial", 
  "Cortes          ", "Resto", 
  "Hecha", "Fecha y Hora", "Piezas",
   "Huella de Cortes"
];

// --- ÍNDICES DE COLUMNAS HISTÓRICO ---
const COL_ID_IDX = 1;
const COL_PERFIL_IDX = 3;
const COL_BARRA_IDX = 4;
const COL_TIPO_MAT_IDX = 5;
const COL_LONG_INI_IDX = 6;
const COL_CORTES_IDX = 7;
const COL_RESTO_IDX = 8;
const COL_HECHA_IDX = 9;
const COL_FECHA_IDX = 10;
const COL_HUELLA_IDX = 11;

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
// Variable global mutable para el backtracking
let horaInicioAlgoritmo;
const HOJA_DATOS_NOMBRE = 'Corredera_GP-60'; 
const HOJA_INVENTARIO_NOMBRE = 'Inventario';
const HOJA_CONFIG_NOMBRE = 'Configuración';
const HOJA_HISTORICO_NOMBRE = 'Plan_Cortes';
const HOJA_MATERIAL_NECESARIO = 'Material_Necesario';

// --- COLORES VISUALES (Inventario) ---
const COLORES_PERFIL = {
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

// Variable global para alertas
let CONFIG = {};
let ALERTAS = [];