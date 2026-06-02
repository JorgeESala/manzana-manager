function mostrarVentanaReporte() {
  const html = `
    <div style="font-family: sans-serif; padding: 15px;">
      <h3 style="color: #333; margin-bottom: 15px;">Generar Reporte Financiero</h3>
      <label style="display:block; font-size:12px; font-weight:bold; margin-bottom:5px;">Fecha Inicio:</label>
      <input type="date" id="fechaInicio" style="width:100%; padding:8px; margin-bottom:15px; border:1px solid #ccc; border-radius:4px;">
      
      <label style="display:block; font-size:12px; font-weight:bold; margin-bottom:5px;">Fecha Fin:</label>
      <input type="date" id="fechaFin" style="width:100%; padding:8px; margin-bottom:20px; border:1px solid #ccc; border-radius:4px;">
      
      <button onclick="ejecutar()" style="width:100%; background:#0f172a; color:white; border:none; padding:10px; font-weight:bold; border-radius:4px; cursor:pointer;">
        Generar Estado de Flujo
      </button>
    </div>
    <script>
      function ejecutar() {
        const fInicio = document.getElementById('fechaInicio').value;
        const fFin = document.getElementById('fechaFin').value;
        if(!fInicio || !fFin) { alert('Por favor selecciona ambas fechas'); return; }
        
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .construirEstadoFinanciero(fInicio, fFin);
      }
    </script>
  `;
  const output = HtmlService.createHtmlOutput(html).setWidth(320).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(output, 'Parámetros del Reporte');
}

function construirEstadoFinanciero(strFechaInicio, strFechaFin) {
  const ssContabilidad = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEgresos = ssContabilidad.getSheetByName("Egresos");
  
  // Convertir strings de los inputs HTML a objetos Date reales (ajustando zona horaria)
  const fechaInicio = new Date(strFechaInicio + "T00:00:00");
  const fechaFin = new Date(strFechaFin + "T23:59:59");

  // 🔗 CONFIGURACIÓN DEL PUENTE CON EL ARCHIVO DE INGRESOS
  const ID_ARCHIVO_INGRESOS = "1MROUyVaih7gAcS1MhCAVt4167Em7ewiugqLOnNvILbY"; 
  const ssIngresos = SpreadsheetApp.openById(ID_ARCHIVO_INGRESOS);
  const sheetHistorial = ssIngresos.getSheetByName("Historial");

  // 1. FILTRAR Y AGRUPAR INGRESOS POR CATEGORÍA EN EL RANGO
  const datosIngresos = sheetHistorial.getDataRange().getValues();
  let sumaMantenimiento = 0;
  let sumaRecargos = 0;

  for (let i = 1; i < datosIngresos.length; i++) {
    const fechaPago = new Date(datosIngresos[i][0]);
    if (fechaPago >= fechaInicio && fechaPago <= fechaFin) {
      const concepto = datosIngresos[i][4];
      const monto = parseFloat(datosIngresos[i][6]) || 0;

      if (concepto === "Mantenimiento") {
        sumaMantenimiento += monto;
      } else if (concepto === "RECARGO") {
        sumaRecargos += monto;
      }
    }
  }

  // 2. FILTRAR Y AGRUPAR EGRESOS POR CATEGORÍA EN EL RANGO
  const datosEgresos = sheetEgresos.getDataRange().getValues();
  const resumenEgresos = {};
  let totalEgresosRango = 0;

  for (let j = 1; j < datosEgresos.length; j++) {
    const fechaGasto = new Date(datosEgresos[j][0]);
    if (fechaGasto >= fechaInicio && fechaGasto <= fechaFin) {
      const categoria = datosEgresos[j][2];
      const monto = parseFloat(datosEgresos[j][4]) || 0;

      if (!resumenEgresos[categoria]) {
        resumenEgresos[categoria] = 0;
      }
      resumenEgresos[categoria] += monto;
      totalEgresosRango += monto;
    }
  }

  // 3. CREAR LA NUEVA HOJA IMPRIMIBLE DEL REPORTE
  const nombreReporte = "Reporte_" + strFechaInicio.replace(/-/g,"") + "_" + strFechaFin.replace(/-/g,"");
  let sheetReporte = ssContabilidad.getSheetByName(nombreReporte);
  if (sheetReporte) ssContabilidad.deleteSheet(sheetReporte);
  sheetReporte = ssContabilidad.insertSheet(nombreReporte);

  // Estilizar Hoja de Reporte (Senior Look)
  sheetReporte.setColumnWidth(1, 280);
  sheetReporte.setColumnWidth(2, 140);
  sheetReporte.getRange("A1:B1").merge().setValue("ESTADO DE FLUJO DE EFECTIVO").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");
  sheetReporte.getRange("A2:B2").merge().setValue("Periodo: " + strFechaInicio + " al " + strFechaFin).setFontStyle("italic").setHorizontalAlignment("center");
  
  let filaActual = 4;

  // --- SECCIÓN INGRESOS ---
  sheetReporte.getRange(filaActual, 1).setValue("INGRESOS (Entradas de Efectivo)").setFontWeight("bold").setFontColor("#1e3a8a");
  filaActual++;
  sheetReporte.getRange(filaActual, 1, 2, 2).setValues([
    ["   Cuotas de Mantenimiento Recaudadas", sumaMantenimiento],
    ["   Cargos por Pago Tardío (Recargos)", sumaRecargos]
  ]);
  filaActual += 2;
  
  const totalIngresosRango = sumaMantenimiento + sumaRecargos;
  sheetReporte.getRange(filaActual, 1, 1, 2).setValues([["TOTAL INGRESOS", totalIngresosRango]]).setFontWeight("bold").setBackground("#f0fdf4");
  filaActual += 2;

  // --- SECCIÓN EGRESOS ---
  sheetReporte.getRange(filaActual, 1).setValue("EGRESOS (Salidas de Efectivo)").setFontWeight("bold").setFontColor("#991b1b");
  filaActual++;
  
  for (const cat in resumenEgresos) {
    sheetReporte.getRange(filaActual, 1, 1, 2).setValues([["   " + cat, resumenEgresos[cat]]]);
    filaActual++;
  }
  
  sheetReporte.getRange(filaActual, 1, 1, 2).setValues([["TOTAL EGRESOS", totalEgresosRango]]).setFontWeight("bold").setBackground("#fef2f2");
  filaActual += 2;

  // --- RENDIMIENTO / FLUJO NETO ---
  const flujoNeto = totalIngresosRango - totalEgresosRango;
  sheetReporte.getRange(filaActual, 1, 1, 2).setValues([["FLUJO NETO DEL PERIODO (Utilidad)", flujoNeto]])
              .setFontWeight("bold")
              .setBackground(flujoNeto >= 0 ? "#bbf7d0" : "#fecaca")
              .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.DOUBLE);

  // Formatear números a Moneda ($)
  sheetReporte.getRange("B4:B" + filaActual).setNumberFormat("$#,##0.00");
  
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheetReporte);
  SpreadsheetApp.getUi().alert("📊 ¡Reporte Generado!\n\nSe ha creado una pestaña llamada '" + nombreReporte + "' con el Estado de Flujo estructurado listo para exportar a PDF o imprimir.");
}