function calcularDineroDisponibleInversion() {
  const ssContabilidad = SpreadsheetApp.getActiveSpreadsheet();

  let sheetEgresos = ssContabilidad.getSheetByName("Egresos");
  if (!sheetEgresos) return;

  const ID_ARCHIVO_INGRESOS = "1MROUyVaih7gAcS1MhCAVt4167Em7ewiugqLOnNvILbY";
  let ssIngresos, sheetHistorial, sheetPropietarios;
  try {
    ssIngresos = SpreadsheetApp.openById(ID_ARCHIVO_INGRESOS);
    sheetHistorial = ssIngresos.getSheetByName("Historial");
    sheetPropietarios = ssIngresos.getSheetByName("Propietarios");
    if (!sheetHistorial || !sheetPropietarios) return;
  } catch (e) { return; }

  const ingresosData = sheetHistorial.getDataRange().getValues();
  let ingresosMZ17 = 0, ingresosMZ19 = 0, ingresosRecargoMZ17 = 0, ingresosRecargoMZ19 = 0;

  for (let i = 1; i < ingresosData.length; i++) {
    if (ingresosData[i][6]) {
      const monto = parseFloat(ingresosData[i][6]);
      const bloque = (ingresosData[i][9] || "").toString().trim();
      const concepto = (ingresosData[i][4] || "").toString().trim();
      if (concepto === "Mantenimiento") {
        if (bloque === "MZ 17") ingresosMZ17 += monto;
        else if (bloque === "MZ 19") ingresosMZ19 += monto;
      } else if (concepto === "RECARGO") {
        if (bloque === "MZ 17") ingresosRecargoMZ17 += monto;
        else if (bloque === "MZ 19") ingresosRecargoMZ19 += monto;
      }
    }
  }

  const egresosData = sheetEgresos.getDataRange().getValues();
  const resumenEgresosMZ17 = {};
  const resumenEgresosMZ19 = {};
  let totalEgresosMZ17 = 0, totalEgresosMZ19 = 0;
  let totalEgresosCompartidoOriginal = 0;
  const detalleCompartido = [];

  for (let j = 1; j < egresosData.length; j++) {
    if (egresosData[j][5]) {
      const monto = parseFloat(egresosData[j][5]);
      const bloque = (egresosData[j][3] || "").toString().trim();
      const categoria = (egresosData[j][2] || "").toString().trim();

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

  let sheetDash = ssContabilidad.getSheetByName("Dashboard");
  if (!sheetDash) {
    sheetDash = ssContabilidad.insertSheet("Dashboard");
  }
  sheetDash.clear();

  sheetDash.setColumnWidth(1, 260);
  sheetDash.setColumnWidth(2, 180);
  sheetDash.setColumnWidth(3, 180);

  const bgHeader = "#1e3a8a";
  const bgSection = "#e0e0e0";
  const bgTotalIngresos = "#f0fdf4";
  const bgTotalEgresos = "#fef2f2";
  const bgDisponible = "#1e3a8a";
  const bgGray = "#f0f0f0";

  let f = 1;
  sheetDash.getRange(f, 1, 1, 3).merge().setValue("CUADRO DE MANDO FINANCIERO").setFontWeight("bold").setFontColor("white").setBackground("#333333").setFontSize(14).setHorizontalAlignment("center");
  f++;

  sheetDash.getRange(f, 1, 1, 3).merge().setValue("RESUMEN GENERAL").setFontWeight("bold").setBackground(bgHeader).setFontColor("white").setHorizontalAlignment("center");
  f++;
  sheetDash.getRange(f, 1).setValue("(+) Total Ingresos Recaudados:");
  sheetDash.getRange(f, 2).setValue(totalIngresos).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 2).setNumberFormat("$#,##0.00");
  f++;
  sheetDash.getRange(f, 1).setValue("(-) Total Egresos Ejecutados:");
  sheetDash.getRange(f, 2).setValue(totalEgresos).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 2).setNumberFormat("$#,##0.00");
  f++;
  sheetDash.getRange(f, 1).setValue("💰 DISPONIBLE TOTAL:");
  sheetDash.getRange(f, 2).setValue(disponibleTotal).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 1, 1, 2).setBackground(disponibleTotal >= 0 ? "#bbf7d0" : "#fecaca");
  sheetDash.getRange(f, 2).setNumberFormat("$#,##0.00");
  f += 2;

  sheetDash.getRange(f, 1).setValue("Coeficientes de reparto (según número de propietarios):");
  sheetDash.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgGray);
  f++;
  sheetDash.getRange(f, 1).setValue("MZ 17: " + countMZ17 + " unidades → " + Math.round(pctMZ17 * 100) + "%");
  sheetDash.getRange(f, 2).setValue("MZ 19: " + countMZ19 + " unidades → " + Math.round(pctMZ19 * 100) + "%");
  f += 2;

  sheetDash.getRange(f, 1, 1, 3).setValues([["Concepto", "MZ 17", "MZ 19"]]).setFontWeight("bold").setBackground(bgHeader).setFontColor("white").setHorizontalAlignment("center");
  f++;
  sheetDash.getRange(f, 1).setValue("INGRESOS");
  sheetDash.getRange(f, 2).setValue("Ingresos MZ 17");
  sheetDash.getRange(f, 3).setValue("Ingresos MZ 19");
  sheetDash.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgSection).setHorizontalAlignment("left");
  f++;
  sheetDash.getRange(f, 1).setValue("   Mantenimiento");
  sheetDash.getRange(f, 2).setValue(ingresosMZ17).setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 3).setValue(ingresosMZ19).setNumberFormat("$#,##0.00");
  f++;
  if (ingresosRecargoMZ17 > 0 || ingresosRecargoMZ19 > 0) {
    sheetDash.getRange(f, 1).setValue("   Recargos por pago tardío");
    sheetDash.getRange(f, 2).setValue(ingresosRecargoMZ17).setNumberFormat("$#,##0.00");
    sheetDash.getRange(f, 3).setValue(ingresosRecargoMZ19).setNumberFormat("$#,##0.00");
    f++;
  }
  sheetDash.getRange(f, 1).setValue("TOTAL INGRESOS");
  sheetDash.getRange(f, 2).setValue(totalIngresosMZ17).setFontWeight("bold");
  sheetDash.getRange(f, 3).setValue(totalIngresosMZ19).setFontWeight("bold");
  sheetDash.getRange(f, 1, 1, 3).setBackground(bgTotalIngresos).setFontWeight("bold");
  sheetDash.getRange(f, 2).setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 3).setNumberFormat("$#,##0.00");
  f += 2;

  sheetDash.getRange(f, 1).setValue("EGRESOS");
  sheetDash.getRange(f, 2).setValue("Egresos MZ 17");
  sheetDash.getRange(f, 3).setValue("Egresos MZ 19");
  sheetDash.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgSection).setHorizontalAlignment("left");
  f++;

  const todasCat = [...new Set([...Object.keys(resumenEgresosMZ17), ...Object.keys(resumenEgresosMZ19)])].sort();
  todasCat.forEach(cat => {
    sheetDash.getRange(f, 1).setValue("   " + cat);
    sheetDash.getRange(f, 2).setValue(resumenEgresosMZ17[cat] || 0).setNumberFormat("$#,##0.00");
    sheetDash.getRange(f, 3).setValue(resumenEgresosMZ19[cat] || 0).setNumberFormat("$#,##0.00");
    f++;
  });

  if (totalEgresosCompartidoOriginal > 0) {
    sheetDash.getRange(f, 1).setValue("   Gastos Compartidos (monto original)");
    sheetDash.getRange(f, 2).setValue(totalEgresosCompartidoOriginal).setFontWeight("bold").setNumberFormat("$#,##0.00");
    sheetDash.getRange(f, 3).setValue(totalEgresosCompartidoOriginal).setFontWeight("bold").setNumberFormat("$#,##0.00");
    f++;
    sheetDash.getRange(f, 1).setValue("      → Proporción MZ 17 (" + Math.round(pctMZ17 * 100) + "%):");
    sheetDash.getRange(f, 2).setValue(sharedMZ17).setNumberFormat("$#,##0.00");
    sheetDash.getRange(f, 1).setFontStyle("italic");
    f++;
    sheetDash.getRange(f, 1).setValue("      → Proporción MZ 19 (" + Math.round(pctMZ19 * 100) + "%):");
    sheetDash.getRange(f, 2).setValue(sharedMZ19).setNumberFormat("$#,##0.00");
    sheetDash.getRange(f, 1).setFontStyle("italic");
    f++;
  }

  const totalGastosMZ17 = totalEgresosMZ17 + sharedMZ17;
  const totalGastosMZ19 = totalEgresosMZ19 + sharedMZ19;
  sheetDash.getRange(f, 1).setValue("TOTAL EGRESOS");
  sheetDash.getRange(f, 2).setValue(totalGastosMZ17).setFontWeight("bold");
  sheetDash.getRange(f, 3).setValue(totalGastosMZ19).setFontWeight("bold");
  sheetDash.getRange(f, 1, 1, 3).setBackground(bgTotalEgresos).setFontWeight("bold");
  sheetDash.getRange(f, 2).setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 3).setNumberFormat("$#,##0.00");
  f += 2;

  sheetDash.getRange(f, 1).setValue("DINERO DISPONIBLE");
  sheetDash.getRange(f, 2).setValue("Disponible MZ 17");
  sheetDash.getRange(f, 3).setValue("Disponible MZ 19");
  sheetDash.getRange(f, 1, 1, 3).setFontWeight("bold").setBackground(bgDisponible).setFontColor("white").setHorizontalAlignment("left").setFontSize(12);
  f++;
  sheetDash.getRange(f, 1).setValue(disponibleMZ17 >= 0 ? "   ✅ Sobró → MZ 17:" : "   ⚠️ Faltó → MZ 17:");
  sheetDash.getRange(f, 2).setValue(disponibleMZ17).setFontWeight("bold").setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 3).setValue(disponibleMZ19 >= 0 ? "✅ Sobró → MZ 19:" : "⚠️ Faltó → MZ 19:").setFontWeight("bold");
  sheetDash.getRange(f, 1, 1, 3).setBackground(disponibleMZ17 >= 0 ? "#bbf7d0" : "#fecaca");
  sheetDash.getRange(f, 3).setValue(disponibleMZ19).setNumberFormat("$#,##0.00");
  sheetDash.getRange(f, 1, 1, 3).setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.THICK);
  f += 2;

  if (detalleCompartido.length > 0) {
    sheetDash.getRange(f, 1).setValue("Detalle de Gastos Compartidos:");
    sheetDash.getRange(f, 1, 1, 3).setFontWeight("bold").setFontColor("#666666").setFontSize(10);
    f++;
    detalleCompartido.forEach(d => {
      sheetDash.getRange(f, 1).setValue("   " + d.desc + ": $" + d.monto.toLocaleString()).setFontColor("#666666").setFontSize(10);
      f++;
    });
  }
}