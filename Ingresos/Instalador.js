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
        ["PROPIETARIO:", "SALDO TOTAL:", "ESTATUS PISCINA"]
      ]
    },
    {
      nombre: "Propietarios",
      color: "#674ea7",
      encabezados: [["ID", "Edificio", "Depto", "Nombre", "Teléfono", "Correo", "Estatus Dispositivo"]]
    },
    {
      nombre: "Cargos",
      color: "#e06666",
      encabezados: [["Fecha Emisión", "Depto", "Concepto", "Monto", "Estatus", "Mes Correspondiente"]]
    },
    {
      nombre: "Historial",
      color: "#38761d",
      encabezados: [["Fecha Pago", "Edificio", "Depto", "Nombre", "Concepto", "Detalle", "Monto", "Forma de Pago"]]
    },
    {
      nombre: "Egresos",
      color: "#f1c232",
      encabezados: [["Fecha", "Categoría", "Proveedor", "Detalle", "Monto"]]
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
      // En lugar de borrar la hoja, la limpiamos completamente para evitar errores de referencia
      hoja.clear();
      hoja.clearFormats();
    } else {
      hoja = ss.insertSheet(h.nombre);
    }
    
    // CORRECCIÓN: Ajuste dinámico del rango según el número de filas y columnas del encabezado
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

}