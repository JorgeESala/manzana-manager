function calcularDineroDisponibleInversion() {
  const ssContabilidad = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEgresos = ssContabilidad.getSheetByName("Egresos");
  
  // 🔗 EL PUENTE: Abre el primer archivo usando su ID único de la URL
  // Copia el ID que sale en la URL de tu hoja de Ingresos (entre /d/ y /edit)
  const ID_ARCHIVO_INGRESOS = "1MROUyVaih7gAcS1MhCAVt4167Em7ewiugqLOnNvILbY"; 
  
  const ssIngresos = SpreadsheetApp.openById(ID_ARCHIVO_INGRESOS);
  const sheetHistorial = ssIngresos.getSheetByName("Historial");
  
  // 1. EXTRAER INGRESOS (Del archivo 1)
  const ingresosData = sheetHistorial.getDataRange().getValues();
  let totalIngresos = 0;
  for (let i = 1; i < ingresosData.length; i++) {
    // Si la fila tiene monto en la columna G (index 6) y no es fila de recargo duplicada
    if (ingresosData[i][6] && ingresosData[i][4] === "Mantenimiento") { 
      totalIngresos += parseFloat(ingresosData[i][6]);
    }
    // Sumar también los recargos cobrados reales (index 6 de las filas RECARGO)
    if (ingresosData[i][6] && ingresosData[i][4] === "RECARGO") {
      totalIngresos += parseFloat(ingresosData[i][6]);
    }
  }
  
  // 2. EXTRAER EGRESOS (De este archivo nuevo)
  const egresosData = sheetEgresos.getDataRange().getValues();
  let totalEgresos = 0;
  for (let j = 1; j < egresosData.length; j++) {
    if (egresosData[j][4]) { // Columna E (Monto del gasto)
      totalEgresos += parseFloat(egresosData[j][4]);
    }
  }
  
  // 3. BALANCE FINAL
  const disponibleInversion = totalIngresos - totalEgresos;
  
  // Renderizar en una pestaña "Dashboard" para que se vea estético en la hoja
  const sheetDash = ssContabilidad.getSheetByName("Dashboard");
  if (sheetDash) {
    sheetDash.getRange("B2").setValue(totalIngresos);
    sheetDash.getRange("B3").setValue(totalEgresos);
    sheetDash.getRange("B4").setValue(disponibleInversion);
  }
  
  SpreadsheetApp.getUi().alert(
    "📊 FLUJO DE CAJA CONSOLIDADO\n\n" +
    "💰 Total Recaudado (Ingresos): $" + totalIngresos.toLocaleString() + "\n" +
    "💸 Total Gastado (Egresos): $" + totalEgresos.toLocaleString() + "\n" +
    "-----------------------------------------\n" +
    "📈 DISPONIBLE PARA INVERSIÓN: $" + disponibleInversion.toLocaleString()
  );
}