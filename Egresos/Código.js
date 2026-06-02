function abrirFormularioEgreso() {
  const html = HtmlService.createHtmlOutputFromFile('FormEgreso')
      .setWidth(450)
      .setHeight(550)
      .setTitle('Registrar Egreso / Gasto');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function procesarEgresoHTML(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEgresos = ss.getSheetByName("Egresos");
  
  const fecha = new Date(datos.fecha + "T12:00:00");
  const monto = parseFloat(datos.monto);
  
  // Folio de auditoría automático (EGR-1001, EGR-1002...)
  const ultimoFolio = sheetEgresos.getLastRow() > 1 ? sheetEgresos.getRange(sheetEgresos.getLastRow(), 2).getValue() : "EGR-1000";
  const numeroFolio = parseInt(ultimoFolio.split("-")[1]) + 1;
  const nuevoFolio = "EGR-" + numeroFolio;

  sheetEgresos.appendRow([
    fecha,
    nuevoFolio,
    datos.categoria,
    datos.descripcion,
    monto,
    datos.metodo,
    datos.folioOperacion || "---"
  ]);

  return "✅ Egreso registrado bajo el folio: " + nuevoFolio;
}
