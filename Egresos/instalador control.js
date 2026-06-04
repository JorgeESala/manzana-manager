function instalarEstructuraContabilidad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Crear Hoja Egresos
  let sheetEgresos = ss.getSheetByName("Egresos") || ss.insertSheet("Egresos");
  sheetEgresos.clear();
  sheetEgresos.getRange("A1:H1").setValues([["Fecha Gasto", "Folio Interno", "Categoría", "Manzana", "Descripción Detallada", "Monto Total", "Método Pago", "Folio Operación / Referencia"]])
             .setBackground("#990000").setFontColor("white").setFontWeight("bold");
  
  // 2. Crear Hoja Dashboard (Indicadores Visuales Financieros)
  let sheetDash = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  sheetDash.clear();
  
  // Diseño del cuadro de control
  sheetDash.getRange("A1:B1").setValues([["CUADRO DE MANDO FINANCIERO", ""]]).setFontWeight("bold").setBackground("#333333").setFontColor("white");
  sheetDash.getRange("A2:B2").setValues([["(+) Total Ingresos Recaudados:", 0]]);
  sheetDash.getRange("A3:B3").setValues([["(-) Total Egresos Ejecutados:", 0]]);
  sheetDash.getRange("A4:B4").setValues([["💰 DISPONIBLE PARA INVERSIÓN:", 0]]).setFontWeight("bold");
  
  // Formato de moneda para los resultados
  sheetDash.getRange("B2:B4").setNumberFormat("$#,##0.00");
  sheetDash.setColumnWidth(1, 250);
  sheetDash.setColumnWidth(2, 150);

  ui.alert("✅ Estructura de CONTABILIDAD y DASHBOARD instalada con éxito.\n\nYa puedes pegar los scripts de registro de egresos y reportes financieros.");
}