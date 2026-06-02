function exportarReporteActualAPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetActual = ss.getActiveSheet();
  const name = sheetActual.getName();
  
  // Validación de seguridad: Solo exportar pestañas que sean reportes
  if (!name.startsWith("Reporte_")) {
    SpreadsheetApp.getUi().alert("⚠️ Operación Cancelada\n\nPor favor, colócate sobre la pestaña del Reporte que deseas exportar antes de presionar este botón.");
    return;
  }
  
  const sheetId = sheetActual.getSheetId();
  const ssId = ss.getId();
  
  // TRUCO SENIOR: Construimos una URL con parámetros de impresión de Google
  // Estos códigos le dicen a Google que lo queremos en vertical, tamaño Carta, ocultando la cuadrícula.
  const url = ss.getUrl().replace(/edit$/, '') + 'export?';
  const exportOptions = {
    exportFormat: 'pdf',
    format: 'pdf',
    size: 'letter',          // Tamaño Carta
    portrait: 'true',        // Orientación Vertical
    fitw: 'true',            // Ajustar al ancho de la página
    gridlines: 'false',      // Ocultar líneas de cuadrícula para estética limpia
    printtitle: 'false',     // Ocultar título del libro
    sheetnames: 'false',     // Ocultar nombre de la hoja
    fzr: 'false',            // No repetir filas congeladas
    gid: sheetId             // Exportar ÚNICAMENTE la pestaña actual
  };
  
  // Juntar todos los parámetros en la URL
  const urlParts = [];
  for (const key in exportOptions) {
    urlParts.push(key + '=' + exportOptions[key]);
  }
  const exportUrl = url + urlParts.join('&');
  
  // INTERFAZ DE DESCARGA: Creamos una ventana HTML temporal con un script de auto-descarga
  const htmlContent = `
    <div style="font-family: sans-serif; text-align: center; padding: 20px;">
      <h4 style="color: #1e3a8a;">📄 Su PDF está listo</h4>
      <p style="font-size: 14px; color: #555;">La descarga comenzará automáticamente en unos segundos.</p>
      <p style="font-size: 12px; color: #999;">Si no inicia, haga clic en el botón de abajo:</p>
      <a href="${exportUrl}" target="_blank" download="${name}.pdf" 
         style="display: inline-block; background: #1e3a8a; color: white; padding: 10px 20px; font-weight: bold; text-decoration: none; border-radius: 4px; margin-top: 10px;">
        Descargar PDF Manual
      </a>
    </div>
    <script>
      // Forzar la apertura/descarga del archivo inmediatamente al abrir la ventana
      window.open("${exportUrl}", "_blank");
      // Cerrar el cuadro de diálogo de Sheets después de 3 segundos
      setTimeout(function() {
        google.script.host.close();
      }, 4000);
    </script>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(350).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generando Archivo...');
}