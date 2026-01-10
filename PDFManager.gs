/**
 * ORQUESTADOR DE PDF: Genera y Guarda en Google Drive
 * Destino: Carpeta "Planes de corte"
 */
function generarPDFTaller() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const NOMBRE_CARPETA_DESTINO = "Planes de corte"; // <--- Tu carpeta destino

  // 1. GENERAR LA HOJA COMBINADA (Visuals.gs)
  try {
    generarImpresionCombinada(); 
  } catch (e) {
    ui.alert("❌ Error generando vista combinada: " + e.message);
    return;
  }

  // 2. PREPARAR NOMBRE DEL ARCHIVO
  const nombreHoja = "Impresion_Combinada";
  const hoja = ss.getSheetByName(nombreHoja);
  
  let nombreProyecto = "PROYECTO_TALLER";
  try {
    // Intentamos leer el título de la celda combinada A1
    const valA1 = hoja.getRange("A1").getValue(); 
    if (valA1 && valA1 !== "") nombreProyecto = valA1;
  } catch(e) {}

  // Añadimos fecha para evitar duplicados exactos
  const fechaStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MM_HHmm");
  const nombreArchivo = `${nombreProyecto}_${fechaStr}.pdf`;

  // 3. GENERAR EL BLOB (El archivo en memoria)
  const blobPDF = crearBlobPDF(ss, nombreHoja, nombreArchivo);
  if (!blobPDF) return;

  // 4. GUARDAR EN DRIVE
  let carpetaDestino;
  const carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_DESTINO);
  
  if (carpetas.hasNext()) {
    carpetaDestino = carpetas.next();
  } else {
    // Si no encuentra la carpeta, avisa y guarda en raíz para no perder el trabajo
    ui.alert(`⚠️ No encontré la carpeta "${NOMBRE_CARPETA_DESTINO}".\nGuardando en "Mi Unidad" (Raíz).`);
    carpetaDestino = DriveApp.getRootFolder();
  }

  const archivoFinal = carpetaDestino.createFile(blobPDF);

  // 5. CONFIRMACIÓN AL USUARIO
  const url = archivoFinal.getUrl();
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <div style="font-family:sans-serif; text-align:center; padding:10px;">
        <p>✅ <strong>Guardado en Drive</strong></p>
        <p style="font-size:12px; color:#666;">Carpeta: ${carpetaDestino.getName()}</p>
        <p style="font-size:14px; font-weight:bold;">${nombreArchivo}</p>
        <br>
        <a href="${url}" target="_blank" style="
           background-color:#1E8E3E; color:white; 
           padding:12px 25px; text-decoration:none; 
           border-radius:5px; font-weight:bold;">
           📂 ABRIR ARCHIVO
        </a>
      </div>`)
    .setWidth(350)
    .setHeight(220);
  ui.showModalDialog(htmlOutput, 'Gestión Documental');
}
/**
 * Genera el BLOB del PDF (Datos crudos) sin guardar archivo todavía
 */
function crearBlobPDF(ss, nombreHoja, nombreArchivoDestino) {
  const hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) return null;

  const ssId = ss.getId();
  const sheetId = hoja.getSheetId();
  const urlBase = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?";
 // Parámetros: A4 HORIZONTAL + AJUSTE A PÁGINA (SCALE=4)
  const params = [
    "exportFormat=pdf", "format=pdf",
    "size=A4", 
    "portrait=false",       // Horizontal
    "scale=4",              // <--- CLAVE: Ajustar a Página (Ancho y Alto)
    "gridlines=false",
    "printtitle=false", "sheetnames=false", "pagenumbers=false",
    // MÁRGENES 0.4 cm (~0.16 in)
    "top_margin=0.16",      
    "bottom_margin=0.16",
    "left_margin=0.16",
    "right_margin=0.16",
    "gid=" + sheetId
  ];
  const urlFinal = urlBase + params.join("&");
  const headers = { "Authorization": "Bearer " + ScriptApp.getOAuthToken() };
  
  try {
    const respuesta = UrlFetchApp.fetch(urlFinal, { headers: headers });
    return respuesta.getBlob().setName(nombreArchivoDestino);
  } catch (e) {
    console.error("Fallo al crear Blob PDF: " + e.message);
    // Si falla, mostramos el error técnico en pantalla para depurar
    SpreadsheetApp.getUi().alert("Error 400 corrigiendo: " + e.message);
    return null;
  }
}