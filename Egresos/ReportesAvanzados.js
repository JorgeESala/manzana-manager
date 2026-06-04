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
  if (!sheetEgresos) return;

  const fechaInicio = new Date(strFechaInicio + "T00:00:00");
  const fechaFin = new Date(strFechaFin + "T23:59:59");

  const ID_ARCHIVO_INGRESOS = "1MROUyVaih7gAcS1MhCAVt4167Em7ewiugqLOnNvILbY";
  let ssIngresos, sheetHistorial, sheetPropietarios;
  try {
    ssIngresos = SpreadsheetApp.openById(ID_ARCHIVO_INGRESOS);
    sheetHistorial = ssIngresos.getSheetByName("Historial");
    sheetPropietarios = ssIngresos.getSheetByName("Propietarios");
    if (!sheetHistorial || !sheetPropietarios) return;
  } catch (e) { return; }

  const datosIngresos = sheetHistorial.getDataRange().getValues();
  let ingresosMZ17 = 0, ingresosMZ19 = 0, ingresosRecargoMZ17 = 0, ingresosRecargoMZ19 = 0;

  for (let i = 1; i < datosIngresos.length; i++) {
    const fechaPago = new Date(datosIngresos[i][0]);
    if (fechaPago >= fechaInicio && fechaPago <= fechaFin) {
      const monto = parseFloat(datosIngresos[i][6]) || 0;
      const bloque = (datosIngresos[i][9] || "").toString().trim();
      const concepto = (datosIngresos[i][4] || "").toString().trim();

      if (concepto === "Mantenimiento") {
        if (bloque === "MZ 17") ingresosMZ17 += monto;
        else if (bloque === "MZ 19") ingresosMZ19 += monto;
      } else if (concepto === "RECARGO") {
        if (bloque === "MZ 17") ingresosRecargoMZ17 += monto;
        else if (bloque === "MZ 19") ingresosRecargoMZ19 += monto;
      }
    }
  }

  const datosEgresos = sheetEgresos.getDataRange().getValues();
  const resumenEgresosMZ17 = {};
  const resumenEgresosMZ19 = {};
  let totalEgresosMZ17 = 0, totalEgresosMZ19 = 0;
  let totalEgresosCompartidoOriginal = 0;
  const detalleCompartido = [];

  for (let j = 1; j < datosEgresos.length; j++) {
    const fechaGasto = new Date(datosEgresos[j][0]);
    if (fechaGasto >= fechaInicio && fechaGasto <= fechaFin) {
      const monto = parseFloat(datosEgresos[j][5]) || 0;
      const bloque = (datosEgresos[j][3] || "").toString().trim();
      const categoria = (datosEgresos[j][2] || "").toString().trim();

      if (bloque === "MZ 17") {
        if (!resumenEgresosMZ17[categoria]) resumenEgresosMZ17[categoria] = 0;
        resumenEgresosMZ17[categoria] += monto;
        totalEgresosMZ17 += monto;
      } else if (bloque === "MZ 19") {
        if (!resumenEgresosMZ19[categoria]) resumenEgresosMZ19[categoria] = 0;
        resumenEgresosMZ19[categoria] += monto;
        totalEgresosMZ19 += monto;
      } else {
        totalEgresosCompartidoOriginal += monto;
        detalleCompartido.push({ desc: categoria, monto: monto });
      }
    }
  }

  const datosProp = sheetPropietarios.getRange("H2:H" + sheetPropietarios.getLastRow()).getValues();
  let countMZ17 = 0, countMZ19 = 0;
  for (let i = 0; i < datosProp.length; i++) {
    const b = (datosProp[i][0] || "").toString().trim();
    if (b === "MZ 17") countMZ17++;
    else if (b === "MZ 19") countMZ19++;
  }
  const totalProp = countMZ17 + countMZ19;
  const pctMZ17 = totalProp > 0 ? (countMZ17 / totalProp) : 0.5;
  const pctMZ19 = totalProp > 0 ? (countMZ19 / totalProp) : 0.5;
  const sharedMZ17 = totalEgresosCompartidoOriginal * pctMZ17;
  const sharedMZ19 = totalEgresosCompartidoOriginal * pctMZ19;

  const totalIngresosMZ17 = ingresosMZ17 + ingresosRecargoMZ17;
  const totalIngresosMZ19 = ingresosMZ19 + ingresosRecargoMZ19;
  const totalIngresos = totalIngresosMZ17 + totalIngresosMZ19;
  const totalEgresos = totalEgresosMZ17 + totalEgresosMZ19 + totalEgresosCompartidoOriginal;
  const disponibleMZ17 = totalIngresosMZ17 - totalEgresosMZ17 - sharedMZ17;
  const disponibleMZ19 = totalIngresosMZ19 - totalEgresosMZ19 - sharedMZ19;
  const disponibleTotal = totalIngresos - totalEgresos;

  const nombreReporte = "Reporte_" + strFechaInicio.replace(/-/g,"") + "_" + strFechaFin.replace(/-/g,"");
  let sheetReporte = ssContabilidad.getSheetByName(nombreReporte);
  if (sheetReporte) ssContabilidad.deleteSheet(sheetReporte);
  sheetReporte = ssContabilidad.insertSheet(nombreReporte);

  sheetReporte.setColumnWidth(1, 260);
  sheetReporte.setColumnWidth(2, 180);
  sheetReporte.setColumnWidth(3, 180);

  const bgHeader = "#1e3a8a";
  const bgSection = "#e0e0e0";
  const bgTotalIngresos = "#f0fdf4";
  const bgTotalEgresos = "#fef2f2";
  const bgDisponible = "#1e3a8a";
  const bgGray = "#f0f0f0";

  let f = 1;
  sheetReporte.getRange(f, 1, 1, 3).merge().setValue("ESTADO DE FLUJO DE EFECTIVO").setFontWeight("bold").setFontSize(16).setFontColor("white").setBackground("#333333").setHorizontalAlignment("center");
  f++;
  sheetReporte.getRange(f, 1, 1, 3).merge().setValue("Del " + strFechaInicio + " al " + strFechaFin).setFontStyle("italic").setFontSize(11).setHorizontalAlignment("center");
  f += 2;

  sheetReporte.getRange(f, 1, 1, 3).merge().setValue("RESUMEN GENERAL").setFontWeight("bold").setBackground(bgHeader).setFontColor("white").setHorizontalAlignment("center");
  f++;
  sheetReporte.getRange(f, 1).setValue("(+) Total Ingresos:");
  sheetReporte.getRange(f, 2).setValue(totalIngresos).setFontWeight("bold").setNumberFormat("$#,##0.00");
  f++;
  sheetReporte.getRange(f, 1).setValue("(-) Total Egresos:");
  sheetReporte.getRange(f, 2).setValue(totalEgresos).setFontWeight("bold").setNumberFormat("$#,##0.00");
  f++;
  sheetReporte.getRange(f, 1).setValue("💰 Dinero Disponible del Periodo:");
  sheetReporte.getRange(f, 2).setValue(disponibleTotal).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 1, 1, 2).setBackground(disponibleTotal >= 0 ? "#bbf7d0" : "#fecaca");
  f += 2;

  sheetReporte.getRange(f, 1).setValue("Coeficientes de reparto (según número de propietarios):");
  sheetReporte.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgGray);
  f++;
  sheetReporte.getRange(f, 1).setValue("MZ 17: " + countMZ17 + " unidades → " + Math.round(pctMZ17 * 100) + "%");
  sheetReporte.getRange(f, 2).setValue("MZ 19: " + countMZ19 + " unidades → " + Math.round(pctMZ19 * 100) + "%");
  f += 2;

  sheetReporte.getRange(f, 1, 1, 3).setValues([["Concepto", "MZ 17", "MZ 19"]]).setFontWeight("bold").setBackground(bgHeader).setFontColor("white").setHorizontalAlignment("center");
  f++;
  sheetReporte.getRange(f, 1).setValue("INGRESOS");
  sheetReporte.getRange(f, 2).setValue("Ingresos MZ 17");
  sheetReporte.getRange(f, 3).setValue("Ingresos MZ 19");
  sheetReporte.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgSection).setHorizontalAlignment("left");
  f++;
  sheetReporte.getRange(f, 1).setValue("   Mantenimiento");
  sheetReporte.getRange(f, 2).setValue(ingresosMZ17).setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 3).setValue(ingresosMZ19).setNumberFormat("$#,##0.00");
  f++;
  if (ingresosRecargoMZ17 > 0 || ingresosRecargoMZ19 > 0) {
    sheetReporte.getRange(f, 1).setValue("   Recargos por pago tardío");
    sheetReporte.getRange(f, 2).setValue(ingresosRecargoMZ17).setNumberFormat("$#,##0.00");
    sheetReporte.getRange(f, 3).setValue(ingresosRecargoMZ19).setNumberFormat("$#,##0.00");
    f++;
  }
  sheetReporte.getRange(f, 1).setValue("TOTAL INGRESOS");
  sheetReporte.getRange(f, 2).setValue(totalIngresosMZ17).setFontWeight("bold");
  sheetReporte.getRange(f, 3).setValue(totalIngresosMZ19).setFontWeight("bold");
  sheetReporte.getRange(f, 1, 1, 3).setBackground(bgTotalIngresos).setFontWeight("bold");
  sheetReporte.getRange(f, 2).setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 3).setNumberFormat("$#,##0.00");
  f += 2;

  sheetReporte.getRange(f, 1).setValue("EGRESOS");
  sheetReporte.getRange(f, 2).setValue("Egresos MZ 17");
  sheetReporte.getRange(f, 3).setValue("Egresos MZ 19");
  sheetReporte.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgSection).setHorizontalAlignment("left");
  f++;

  const todasCat = [...new Set([...Object.keys(resumenEgresosMZ17), ...Object.keys(resumenEgresosMZ19)])].sort();
  todasCat.forEach(cat => {
    sheetReporte.getRange(f, 1).setValue("   " + cat);
    sheetReporte.getRange(f, 2).setValue(resumenEgresosMZ17[cat] || 0).setNumberFormat("$#,##0.00");
    sheetReporte.getRange(f, 3).setValue(resumenEgresosMZ19[cat] || 0).setNumberFormat("$#,##0.00");
    f++;
  });

  if (totalEgresosCompartidoOriginal > 0) {
    sheetReporte.getRange(f, 1).setValue("   Gastos Compartidos (monto original)");
    sheetReporte.getRange(f, 2).setValue(totalEgresosCompartidoOriginal).setFontWeight("bold").setNumberFormat("$#,##0.00");
    sheetReporte.getRange(f, 3).setValue(totalEgresosCompartidoOriginal).setFontWeight("bold").setNumberFormat("$#,##0.00");
    f++;
    sheetReporte.getRange(f, 1).setValue("      → Proporción MZ 17 (" + Math.round(pctMZ17 * 100) + "%):");
    sheetReporte.getRange(f, 2).setValue(sharedMZ17).setNumberFormat("$#,##0.00");
    sheetReporte.getRange(f, 1).setFontStyle("italic");
    f++;
    sheetReporte.getRange(f, 1).setValue("      → Proporción MZ 19 (" + Math.round(pctMZ19 * 100) + "%):");
    sheetReporte.getRange(f, 2).setValue(sharedMZ19).setNumberFormat("$#,##0.00");
    sheetReporte.getRange(f, 1).setFontStyle("italic");
    f++;
  }

  const totalGastosMZ17 = totalEgresosMZ17 + sharedMZ17;
  const totalGastosMZ19 = totalEgresosMZ19 + sharedMZ19;
  sheetReporte.getRange(f, 1).setValue("TOTAL EGRESOS");
  sheetReporte.getRange(f, 2).setValue(totalGastosMZ17).setFontWeight("bold");
  sheetReporte.getRange(f, 3).setValue(totalGastosMZ19).setFontWeight("bold");
  sheetReporte.getRange(f, 1, 1, 3).setBackground(bgTotalEgresos).setFontWeight("bold");
  sheetReporte.getRange(f, 2).setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 3).setNumberFormat("$#,##0.00");
  f += 2;

  sheetReporte.getRange(f, 1).setValue("DINERO DISPONIBLE");
  sheetReporte.getRange(f, 2).setValue("Disponible MZ 17");
  sheetReporte.getRange(f, 3).setValue("Disponible MZ 19");
  sheetReporte.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgDisponible).setFontColor("white").setHorizontalAlignment("left").setFontSize(12);
  f++;
  sheetReporte.getRange(f, 1).setValue(disponibleMZ17 >= 0 ? "   ✅ Sobró → MZ 17:" : "   ⚠️ Faltó → MZ 17:");
  sheetReporte.getRange(f, 2).setValue(disponibleMZ17).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 3).setValue(disponibleMZ19 >= 0 ? "✅ Sobró → MZ 19:" : "⚠️ Faltó → MZ 19:").setFontWeight("bold");
  sheetReporte.getRange(f, 1, 1, 3).setBackground(disponibleMZ17 >= 0 ? "#bbf7d0" : "#fecaca");
  sheetReporte.getRange(f, 3).setValue(disponibleMZ19).setNumberFormat("$#,##0.00");
  sheetReporte.getRange(f, 1, 1, 3).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.THICK);
  f += 2;

  if (detalleCompartido.length > 0) {
    sheetReporte.getRange(f, 1).setValue("Detalle de Gastos Compartidos:");
    sheetReporte.getRange(f, 1, 1, 3).setFontWeight("bold").setFontColor("#666666").setFontSize(10);
    f++;
    detalleCompartido.forEach(d => {
      sheetReporte.getRange(f, 1).setValue("   " + d.desc + ": $" + d.monto.toLocaleString()).setFontColor("#666666").setFontSize(10);
      f++;
    });
  }

  ssContabilidad.setActiveSheet(sheetReporte);
}