function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('💼 Administración')
      .addItem('📝 Registrar Nuevo Gasto', 'abrirFormularioEgreso')
      .addSeparator()
      .addItem('📊 Calcular Dinero Disponible', 'calcularDineroDisponibleInversion')
      .addItem('📅 Generar Estado Financiero por Fechas', 'mostrarVentanaReporte')
      .addItem('🖨️ Exportar Reporte Actual a PDF', 'exportarReporteActualAPDF')
      .addToUi();
  } catch (e) {
    // Si la interfaz no está lista, guarda el registro silenciosamente sin romper el script
    Logger.log("La interfaz visual no estaba lista aún: " + e.message);
  }
}