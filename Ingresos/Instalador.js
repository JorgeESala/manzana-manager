function instalarSistemaDesdeCero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Definición de la estructura maestra con tamaños específicos
  const hojas = [
    {
      nombre: "Panel",
      color: "#4a86e8",
      encabezados: [
        ["ADMINISTRADOR MANZANA SEGURA", "", ""],
        ["BUSCADOR DE PROPIETARIO", "", ""],
        ["Ingrese Depto:", "", "Estado de Cuenta"],
        ["PROPIETARIO:", "SALDO TOTAL:", "ESTATUS PISCINA"],
        ["BLOQUE ACTIVO:", "MZ 17", ""]
      ]
    },
    {
      nombre: "Propietarios",
      color: "#674ea7",
      encabezados: [["ID", "Edificio", "Depto", "Nombre", "Teléfono", "Correo", "ID Vivienda", "Manzana"]]
    },
    {
      nombre: "Cargos",
      color: "#e06666",
      encabezados: [["Fecha Emisión", "Depto", "Concepto", "Monto", "Estatus", "Mes Correspondiente", "Manzana", "Saldo"]]
    },
    {
      nombre: "Historial",
      color: "#38761d",
      encabezados: [["Fecha Pago", "Edificio", "Depto", "Nombre", "Concepto", "Detalle", "Monto", "Forma de Pago", "Referencia", "Manzana"]]
    },
    {
      nombre: "Egresos",
      color: "#f1c232",
      encabezados: [["Fecha", "Folio Interno", "Categoría", "Manzana", "Descripción Detallada", "Monto Total", "Método Pago", "Folio Operación / Referencia"]]
    },
    {
      nombre: "Recibos",
      color: "#ff9900",
      encabezados: [
        ["GENERADOR DE RECIBO", "", "", "", "", "FECHA:", ""],
        ["", "", "", "", "", "MES A PAGAR:", ""],
        ["DEPARTAMENTO:", "", "", "", "", "PROPIETARIO:", ""],
        ["", "", "", "", "", "EMAIL:", ""],
        ["", "", "", "", "", "PAGO:", ""],
        ["ID CUOTA", "CONCEPTO", "DETALLE", "MONTO", "", "", ""]
      ]
    }
  ];

  hojas.forEach(h => {
    let hoja = ss.getSheetByName(h.nombre);
    if (hoja) {
      hoja.clear();
      hoja.clearFormats();
    } else {
      hoja = ss.insertSheet(h.nombre);
    }
    
    const filas = h.encabezados.length;
    const columnas = h.encabezados[0].length;
    
    const range = hoja.getRange(1, 1, filas, columnas);
    range.setValues(h.encabezados)
         .setFontWeight("bold")
         .setBackground(h.color)
         .setFontColor("white")
         .setBorder(true, true, true, true, true, true);
         
    hoja.setTabColor(h.color);
    hoja.autoResizeColumns(1, columnas);
  });

  // Add data validation for block selector in Panel
  const panel = ss.getSheetByName("Panel");
  if (panel) {
    const bloqueCell = panel.getRange("B5");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["MZ 17", "MZ 19"], true)
      .setAllowInvalid(false)
      .build();
    bloqueCell.setDataValidation(rule);
    bloqueCell.setFontWeight("bold").setBackground("#e8f0fe");
  }

  // Add data validation for Bloque column in Propietarios (column I)
  const propietarios = ss.getSheetByName("Propietarios");
  if (propietarios) {
    const bloqueRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["MZ 17", "MZ 19"], true)
      .setAllowInvalid(false)
      .build();
    propietarios.getRange("I2:I1000").setDataValidation(bloqueRule);
  }
}
