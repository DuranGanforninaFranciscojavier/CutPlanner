
/**
 * ORQUESTADOR: Regenera la vista visual y saca el PDF automáticamente.
 */
function generarPDFTaller() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. Ordena a Visuals que actualice los datos (tu código existente)
  try {
    generarImpresionCombinada(); 
  } catch (e) {
    ui.alert("❌ Fallo al generar la hoja visual: " + e.message);
    return;
  }

  // 2. Ordena la exportación a PDF (el código que acabas de probar)
  const nombreHoja = "Impresion_Combinada";
  const archivoPDF = exportarHojaAPDF(nombreHoja); // Usamos la función auxiliar que ya tienes
  
  // 3. Entrega el resultado en bandeja de plata
  if (archivoPDF) {
    const url = archivoPDF.getUrl();
    const htmlOutput = HtmlService
      .createHtmlOutput(`<p style="font-family:sans-serif;">Tu orden ha sido ejecutada.</p>
                         <p><a href="${url}" target="_blank" style="background-color:#4CAF50; color:white; padding:10px 20px; text-decoration:none; border-radius:5px;">📂 ABRIR PDF AHORA</a></p>`)
      .setWidth(300)
      .setHeight(150);
    ui.showModalDialog(htmlOutput, '✅ PDF Generado');
  }
}

function exportarHojaAPDF(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(nombreHoja);
  
  if (!hoja) {
    console.error("❌ ERROR: No encuentro la hoja '" + nombreHoja + "'. Ejecuta primero la generación visual.");
    return null;
  }

  const ssId = ss.getId();
  const sheetId = hoja.getSheetId();
  
  // Configuración de la URL de exportación (A4, Vertical, Ajustado al ancho)
  const urlBase = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?";
  const params = [
    "exportFormat=pdf",
    "format=pdf",
    "size=A4",             // Tamaño papel
    "portrait=true",       // Vertical (false para horizontal)
    "fitw=true",           // Ajustar al ancho (CRÍTICO para que no se corte)
    "gridlines=false",     // Sin líneas de cuadrícula
    "printtitle=false",    // Sin títulos
    "sheetnames=false",    // Sin nombre de hoja
    "pagenumbers=false",   // Sin números de página
    "gid=" + sheetId       // ID específico de la hoja
  ];
  
  const urlFinal = urlBase + params.join("&");
  
  const token = ScriptApp.getOAuthToken();
  const headers = {
    "Authorization": "Bearer " + token
  };
  
  console.log("🔄 Generando Blob del PDF...");
  const respuesta = UrlFetchApp.fetch(urlFinal, { headers: headers });
  const blob = respuesta.getBlob().setName(nombreHoja + "_" + new Date().getTime() + ".pdf");
  
  console.log("💾 Guardando en Drive...");
  const archivo = DriveApp.createFile(blob);
  return archivo;
}