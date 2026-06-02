/**
 * PROYECTO: AT-Version1
 * Automatización para MANZANA 80 ALDEA TULUM A.C.
 */

const FOLDER_ID = "1U2r8nB7DHfBIMnfzl366AI0j7VLE4jUP";

// ==========================================
// 1. BOTÓN HOJA "RECIBOS": GUARDAR, PDF Y CORREO (GMAIL WEB)
// ==========================================
function generarRecibo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRecibos = ss.getSheetByName("Recibos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  // 1. CAPTURA DE DATOS DEL FORMULARIO
  const idVivienda = sheetRecibos.getRange("B3").getValue(); // Ej: "P2 101"
  const mesAPagar = sheetRecibos.getRange("G2").getValue();
  const montoPagado = sheetRecibos.getRange("F18").getValue();
  const formaPago = sheetRecibos.getRange("G5").getValue();
  const referencia = sheetRecibos.getRange("G6").getValue();

  if (!idVivienda || montoPagado <= 0) {
    SpreadsheetApp.getUi().alert("⚠️ Selecciona un departamento y un monto válido.");
    return;
  }

  // 2. BUSCAR EL NOMBRE DIRECTAMENTE EN "PROPIETARIOS" (Columna D)
  // Buscamos el ID Vivienda en la Columna H y traemos lo que está en la D
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreReal = "Nombre no encontrado";

  for (let i = 0; i < datosProp.length; i++) {
    // datosProp[i][4] es la columna H (ID Vivienda)
    // datosProp[i][0] es la columna D (Nombre)
    if (datosProp[i][4] === idVivienda) {
      nombreReal = datosProp[i][0];
      break;
    }
  }

  // 3. ACTUALIZAR ESTATUS EN "CARGOS"
  const datosCargos = sheetCargos.getDataRange().getValues();
  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1] === idVivienda && datosCargos[i][5] === mesAPagar && datosCargos[i][4] === "Pendiente") {
      sheetCargos.getRange(i + 1, 5).setValue("Pagado");
      break;
    }
  }

  // 4. LÓGICA DE FECHAS Y RECARGO
  const fechaHoy = new Date();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  let recargo = (mesAPagar !== mesActualStr || fechaHoy.getDate() > 10) ? montoPagado * 0.10 : 0;

  // 5. PASAR AL HISTORIAL
  const partes = idVivienda.split(" ");
  const edificio = partes[0];
  const depto = partes[1] || "";

  // Registro principal
  sheetHistorial.appendRow([
    fechaHoy,
    edificio,
    depto,
    nombreReal, // <--- Aquí ya va el nombre recuperado directamente
    "Mantenimiento",
    "Pago de " + mesAPagar,
    montoPagado,
    formaPago,
    referencia // Columna I
  ]);

  // Registro de recargo si aplica
  if (recargo > 0) {
    sheetHistorial.appendRow([
      fechaHoy,
      edificio,
      depto,
      nombreReal,
      "RECARGO",
      "Recargo 10% " + mesAPagar,
      recargo,
      "Automático",
      "---"
    ]);
  }

  ss.toast("✅ Pago de " + nombreReal + " registrado correctamente.");
  limpiarRecibo();
}

function generarReportePiscina() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const ui = SpreadsheetApp.getUi();

  const datosCargos = sheetCargos.getDataRange().getValues();
  const registroDeudas = {}; // Objeto para agrupar deudas por ID

  // 1. Contabilizar meses pendientes por cada ID Vivienda
  for (let i = 1; i < datosCargos.length; i++) {
    const idVivienda = datosCargos[i][1];
    const estatus = datosCargos[i][4];

    if (estatus === "Pendiente") {
      if (!registroDeudas[idVivienda]) {
        registroDeudas[idVivienda] = {
          meses: [],
          total: 0
        };
      }
      registroDeudas[idVivienda].meses.push(datosCargos[i][5]); // Guarda el nombre del mes
      registroDeudas[idVivienda].total += parseFloat(datosCargos[i][3]);
    }
  }

  // 2. Filtrar solo los que deben 2 o más meses y buscar sus datos de contacto
  const listaBloqueo = [];
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();

  for (const id in registroDeudas) {
    if (registroDeudas[id].meses.length >= 2) {
      // Buscar el nombre y teléfono en la hoja Propietarios
      let nombre = "No encontrado";
      let telefono = "---";

      for (let j = 0; j < datosProp.length; j++) {
        if (datosProp[j][4] === id) { // Columna H
          nombre = datosProp[j][0];   // Columna D
          // Asumiendo que el teléfono está en la columna E de Propietarios (ajusta el index si es otra)
          // telefono = datosProp[j][1]; 
          break;
        }
      }

      listaBloqueo.push([
        id,
        nombre,
        registroDeudas[id].meses.length,
        "$" + registroDeudas[id].total,
        registroDeudas[id].meses.join(", ")
      ]);
    }
  }

  // 3. Mostrar el resultado
  if (listaBloqueo.length === 0) {
    ui.alert("✅ ¡Excelente noticia!\n\nNo hay ninguna vivienda que deba 2 o más meses. Todos pueden usar la piscina.");
  } else {
    // Creamos una tabla simple en HTML para mostrar la lista
    let htmlContent = `<div style="font-family: sans-serif;">
      <p>Las siguientes viviendas tienen <b>2 o más meses de adeudo</b> y su acceso a la piscina debe ser suspendido:</p>
      <table border="1" style="border-collapse: collapse; width: 100%; text-align: left;">
        <tr style="background-color: #f2f2f2;">
          <th style="padding: 8px;">ID Vivienda</th>
          <th style="padding: 8px;">Propietario</th>
          <th style="padding: 8px;">Meses</th>
          <th style="padding: 8px;">Adeudo</th>
        </tr>`;

    listaBloqueo.forEach(row => {
      htmlContent += `<tr>
        <td style="padding: 8px;">${row[0]}</td>
        <td style="padding: 8px;">${row[1]}</td>
        <td style="padding: 8px; color: red; font-weight: bold;">${row[2]}</td>
        <td style="padding: 8px;">${row[3]}</td>
      </tr>`;
    });

    htmlContent += `</table><br><p style="font-size: 12px; color: gray;">Meses adeudados: ${listaBloqueo.map(r => r[0] + " (" + r[4] + ")").join(" | ")}</p></div>`;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(450)
      .setTitle('LISTA DE BLOQUEO - PISCINA');
    ui.showModalDialog(htmlOutput, ' ');
  }
}
function limpiarRecibo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRecibos = ss.getSheetByName("Recibos");

  // NO borramos G3 porque tiene la fórmula. 
  // Al borrar B3 (el buscador), G3 se limpia solo por la fórmula BUSCARV/XLOOKUP.
  sheetRecibos.getRange("B3").clearContent();   // Buscador de ID Vivienda
  sheetRecibos.getRange("G2").clearContent();   // Mes a pagar
  sheetRecibos.getRange("F18:F23").clearContent(); // Montos
  sheetRecibos.getRange("G5:G6").clearContent();   // Forma de pago y Referencia

  ss.toast("Formulario listo para nueva entrada.", "Recibo Limpio");
}

function actualizarEstatusCargo(depa, concepto, mes) {
  const sheetCargos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cargos");
  const datos = sheetCargos.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    // Columna B: Depto, Columna C: Concepto, Columna F: Mes, Columna E: Estatus
    if (datos[i][1] == depa && datos[i][2] == concepto && datos[i][5] == mes) {
      sheetCargos.getRange(i + 1, 5).setValue("Pagado");
      break;
    }
  }
}


// ==========================================
// NUEVO BOTÓN: LIMPIAR RECIBO (Por si te equivocas)
// ==========================================



// ==========================================
// 3. BOTÓN HOJA "CARGOS": VACIAR DE J4:J9 A LA TABLA
// ==========================================
function agregarCargo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Pedir datos mediante ventanas emergentes (más fácil para el usuario)
  const depto = ui.prompt("Nuevo Cargo", "Ingrese el número de departamento:", ui.ButtonSet.OK_CANCEL);
  if (depto.getSelectedButton() != ui.Button.OK) return;

  const monto = ui.prompt("Nuevo Cargo", "Ingrese el monto (ej. 500):", ui.ButtonSet.OK_CANCEL);
  const mes = ui.prompt("Nuevo Cargo", "Ingrese el mes (ej. Mayo):", ui.ButtonSet.OK_CANCEL);

  const sheetCargos = ss.getSheetByName("Cargos");
  sheetCargos.appendRow([new Date(), depto.getResponseText(), "Cuota Mantenimiento", monto.getResponseText(), "Pendiente", mes.getResponseText()]);

  ss.toast("Cargo registrado para el departamento " + depto.getResponseText());
}
function agregarCargoIndividual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Pedir el ID Vivienda (Edificio Depto)
  const deptoRes = ui.prompt("Nuevo Cargo Individual", "Ingrese el ID Vivienda (Ej: P2 101):", ui.ButtonSet.OK_CANCEL);
  if (deptoRes.getSelectedButton() != ui.Button.OK) return;
  const idVivienda = deptoRes.getResponseText();

  // 2. Pedir el Motivo/Concepto
  const motivoRes = ui.prompt("Concepto", "Ingrese el motivo del cargo (ej. Multa, Reparación, Agua):", ui.ButtonSet.OK_CANCEL);
  if (motivoRes.getSelectedButton() != ui.Button.OK) return;
  const motivo = motivoRes.getResponseText();

  // 3. Pedir el Monto
  const montoRes = ui.prompt("Monto", "Ingrese el monto (ej. 500):", ui.ButtonSet.OK_CANCEL);
  if (montoRes.getSelectedButton() != ui.Button.OK) return;
  const monto = montoRes.getResponseText();

  // 4. Pedir el Mes
  const mesRes = ui.prompt("Mes", "Ingrese el mes al que corresponde (ej. Mayo):", ui.ButtonSet.OK_CANCEL);
  if (mesRes.getSelectedButton() != ui.Button.OK) return;
  const mes = mesRes.getResponseText();

  const sheetCargos = ss.getSheetByName("Cargos");

  // Registrar en la hoja de Cargos
  // Formato: [Fecha Emisión, Depto/ID, Concepto, Monto, Estatus, Mes Correspondiente]
  sheetCargos.appendRow([
    new Date(),
    idVivienda,
    motivo,
    monto,
    "Pendiente",
    mes
  ]);

  ss.toast("✅ Cargo de '" + motivo + "' registrado para " + idVivienda);
}
function generarCargosMes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const sheetCargos = ss.getSheetByName("Cargos");
  const ui = SpreadsheetApp.getUi();

  // Preguntar qué mes vamos a cargar
  const respuesta = ui.prompt("Generar Cargos Mensuales", "Escriba el mes (ej. Mayo):", ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  const mes = respuesta.getResponseText();

  const montoRes = ui.prompt("Monto", "Monto de la cuota (ej. 500):", ui.ButtonSet.OK_CANCEL);
  const monto = montoRes.getResponseText();

  const propietarios = sheetPropietarios.getRange("C2:C" + sheetPropietarios.getLastRow()).getValues();
  const fechaActual = new Date();

  // Insertar cargos para todos los que tengan departamento registrado
  propietarios.forEach(fila => {
    if (fila[0] !== "") {
      sheetCargos.appendRow([fechaActual, fila[0], "Mantenimiento", monto, "Pendiente", mes]);
    }
  });

  ui.alert("✅ Se han generado los cargos de mantenimiento para " + propietarios.length + " propietarios.");
}
function registrarPagoInteligente() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetHistorial = ss.getSheetByName("Historial");

  // 1. Pedir Datos Básicos
  const idVivienda = ui.prompt("Pago", "ID Vivienda (Ej: P2 101):", ui.ButtonSet.OK).getResponseText().toUpperCase();
  const mes = ui.prompt("Pago", "Mes que está pagando (Ej: Mayo):", ui.ButtonSet.OK).getResponseText();
  const montoPagado = parseFloat(ui.prompt("Pago", "Monto que transfirió:", ui.ButtonSet.OK).getResponseText());

  // 2. Pedir Fecha (Simulando Date Picker con texto, GAS no tiene picker nativo en prompts)
  const fechaHoy = new Date();
  const fechaSugerida = Utilities.formatDate(fechaHoy, "GMT-6", "dd/MM/yyyy");
  const fechaInput = ui.prompt("Fecha de Pago", "Confirme fecha (DD/MM/YYYY) o deje igual para hoy:", ui.ButtonSet.OK).getResponseText() || fechaSugerida;

  // Convertir input a objeto Date
  const partesFecha = fechaInput.split("/");
  const fechaFinal = new Date(partesFecha[2], partesFecha[1] - 1, partesFecha[0]);

  // 3. Buscar el Cargo en la hoja
  let datos = sheetCargos.getDataRange().getValues();
  let filaEncontrada = -1;
  let montoOriginal = 0;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === idVivienda && datos[i][5] === mes && datos[i][4] === "Pendiente") {
      filaEncontrada = i + 1;
      montoOriginal = datos[i][3];
      break;
    }
  }

  if (filaEncontrada === -1) {
    ui.alert("No se encontró un cargo pendiente para " + idVivienda + " en " + mes);
    return;
  }

  // 4. Lógica de Recargo (Día 10)
  let montoConRecargo = montoOriginal;
  if (fechaFinal.getDate() > 10) {
    // Si no se ha aplicado recargo antes (asumiendo base 500)
    if (montoOriginal === 500) {
      montoConRecargo = 550;
      ss.toast("Se detectó pago extemporáneo. Total actualizado a $550", "Aviso");
    }
  }

  // 5. Procesar el Pago
  const saldoRestante = montoConRecargo - montoPagado;

  if (saldoRestante <= 0) {
    // Pago completo o de más
    sheetCargos.getRange(filaEncontrada, 5).setValue("Pagado");
    sheetCargos.getRange(filaEncontrada, 4).setValue(montoConRecargo); // Actualiza si hubo recargo
    ui.alert("✅ Pago completo registrado.");
  } else {
    // Pago parcial
    sheetCargos.getRange(filaEncontrada, 4).setValue(saldoRestante);
    ui.alert("⚠️ Pago parcial. El cargo queda pendiente por: $" + saldoRestante);
  }

  // 6. Registrar en Historial
  sheetHistorial.appendRow([fechaFinal, idVivienda.split(" ")[0], idVivienda.split(" ")[1], "---", "Mantenimiento", "Abono " + mes, montoPagado, "Transferencia/Efectivo"]);
}
function generarCargosMesPro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProp = ss.getSheetByName("Propietarios");
  const sheetCargos = ss.getSheetByName("Cargos");
  const ui = SpreadsheetApp.getUi();

  const mes = ui.prompt("Mes de Cargo", "Ingrese el mes (ej. Mayo):", ui.ButtonSet.OK).getResponseText();
  const monto = ui.prompt("Monto", "Monto de cuota:", ui.ButtonSet.OK).getResponseText();

  // Obtenemos Edificio (Col B) y Depto (Col C)
  const datos = sheetProp.getRange("B2:C" + sheetProp.getLastRow()).getValues();
  const fecha = new Date();

  datos.forEach(fila => {
    if (fila[0] !== "" && fila[1] !== "") {
      const idVivienda = fila[0] + " " + fila[1]; // Unión: "P2 101"

      // Antes de agregar, verificamos si ya existe el cargo para no duplicar
      sheetCargos.appendRow([fecha, idVivienda, "Mantenimiento", monto, "Pendiente", mes]);
    }
  });

  ui.alert("Cargos generados exitosamente para todas las unidades.");
}
function generarListaBloqueo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const datos = sheetCargos.getDataRange().getValues();
  const deudores = {};

  // Empezamos en i=1 para saltar el encabezado
  for (let i = 1; i < datos.length; i++) {
    let depa = datos[i][1];
    let estatus = datos[i][4]; // Columna E: Estatus

    if (estatus === "Pendiente") {
      deudores[depa] = (deudores[depa] || 0) + 1;
    }
  }

  let listaBloqueo = [];
  for (let depa in deudores) {
    if (deudores[depa] >= 2) {
      listaBloqueo.push([depa, deudores[depa]]);
    }
  }

  if (listaBloqueo.length > 0) {
    let msj = "Propietarios con 2 o más meses de adeudo:\n\n";
    listaBloqueo.forEach(res => {
      msj += "Depto: " + res[0] + " - Meses pendientes: " + res[1] + "\n";
    });
    SpreadsheetApp.getUi().alert("LISTA DE BLOQUEO PISCINA", msj, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("No hay deudores críticos para bloqueo.");
  }
}

// Servir la página web al entrar a la URL
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Portal')
    .setTitle("Buscador De Estados de Cuenta")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Función que consultará el propietario desde su celular
function obtenerEstadoCuenta(idVivienda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  const idBusqueda = idVivienda.toUpperCase().trim();
  const fechaHoy = new Date();
  const diaActual = fechaHoy.getDate();

  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  // 1. Buscar Nombre
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreEncontrado = "";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][4].toString().toUpperCase().trim() === idBusqueda) {
      nombreEncontrado = datosProp[i][0];
      break;
    }
  }
  if (!nombreEncontrado) return { error: true };

  // 2. Buscar Adeudos y calcular recargos "al vuelo"
  const cargos = sheetCargos.getDataRange().getValues();
  let totalAdeudo = 0;
  let desglose = [];
  let contadorMesesPendientes = 0;

  for (let i = 1; i < cargos.length; i++) {
    if (cargos[i][1].toString().toUpperCase().trim() === idBusqueda && cargos[i][4] === "Pendiente") {
      let montoCargo = parseFloat(cargos[i][3]);
      let mesCargo = cargos[i][5];
      let concepto = cargos[i][2];

      // APLICAR RECARGO EN TIEMPO REAL PARA EL PORTAL
      // Si es mantenimiento, no se ha cobrado recargo aún (monto es 500) 
      // Y ya pasó del día 10 del mes actual O es un mes anterior atrasado.
      if (concepto === "Mantenimiento" && montoCargo === 500) {
        if (mesCargo !== mesActualStr || diaActual > 10) {
          montoCargo += 50; // Sumamos los $50 de recargo virtualmente para la pantalla
        }
      }

      totalAdeudo += montoCargo;
      desglose.push({
        mes: mesCargo + (montoCargo > 500 ? " (+Recargo)" : ""),
        monto: montoCargo
      });
      contadorMesesPendientes++;
    }
  }

  return {
    nombre: nombreEncontrado,
    saldo: totalAdeudo,
    piscina: contadorMesesPendientes >= 2 ? "BLOQUEADO" : "ACCESO PERMITIDO",
    detalles: desglose
  };
}

function busquedaGlobalPropietario() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Pedir el ID (Edificio + Depto)
  const respuesta = ui.prompt("Búsqueda Global", "Ingrese Edificio y Depto (Ej: P2 101):", ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  const idBusqueda = respuesta.getResponseText().toUpperCase();

  const datosCargos = ss.getSheetByName("Cargos").getDataRange().getValues();
  let tablaHtml = "<table border='1' style='width:100%; border-collapse:collapse; font-family:Arial;'>";
  tablaHtml += "<tr style='background-color:#4a86e8; color:white;'><th>Mes</th><th>Concepto</th><th>Monto</th><th>Estatus</th></tr>";

  let encontrado = false;
  let deudaTotal = 0;

  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1].toString().toUpperCase() === idBusqueda) {
      const mes = datosCargos[i][5];
      const concepto = datosCargos[i][2];
      const monto = datosCargos[i][3];
      const estatus = datosCargos[i][4];

      const colorEstatus = estatus === "Pendiente" ? "red" : "green";
      if (estatus === "Pendiente") deudaTotal += monto;

      tablaHtml += `<tr><td>${mes}</td><td>${concepto}</td><td>$${monto}</td><td style="color:${colorEstatus}; font-weight:bold;">${estatus}</td></tr>`;
      encontrado = true;
    }
  }
  tablaHtml += "</table>";

  if (!encontrado) {
    ui.alert("No se encontró información para: " + idBusqueda);
    return;
  }

  const headerHtml = `<h3>Resumen para: ${idBusqueda}</h3><p><b>Deuda Pendiente Total: $${deudaTotal}</b></p>`;
  const htmlOutput = HtmlService.createHtmlOutput(headerHtml + tablaHtml)
    .setWidth(450)
    .setHeight(400);

  ui.showModalDialog(htmlOutput, "Estado de Cuenta Detallado");
}

function mostrarEstadoCuentaCompleto() {
  const html = `
    <div style="font-family:sans-serif;padding:20px;">
      <h3 style="margin:0 0 15px 0;color:#1e3a8a;">Estado de Cuenta</h3>
      <div class="mb-3">
        <label style="display:block;font-size:12px;font-weight:bold;margin-bottom:4px;color:#333;">ID Vivienda</label>
        <input type="text" id="inputId" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;" placeholder="Ej: P2 101">
      </div>
      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <div style="flex:1;">
          <label style="display:block;font-size:12px;font-weight:bold;margin-bottom:4px;color:#333;">Fecha Inicio</label>
          <input type="date" id="inputFechaInicio" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">
        </div>
        <div style="flex:1;">
          <label style="display:block;font-size:12px;font-weight:bold;margin-bottom:4px;color:#333;">Fecha Fin</label>
          <input type="date" id="inputFechaFin" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">
        </div>
      </div>
      <button onclick="generar()" style="width:100%;background:#1e3a8a;color:white;border:none;padding:10px;font-weight:bold;border-radius:4px;cursor:pointer;">GENERAR ESTADO DE CUENTA</button>
      <div id="msg" style="margin-top:10px;font-size:13px;color:#666;text-align:center;"></div>
    </div>
    <script>
      document.getElementById('inputFechaInicio').valueAsDate = new Date(new Date().getFullYear(), new Date().getMonth(), 1);
      document.getElementById('inputFechaFin').valueAsDate = new Date();
      function generar() {
        const id = document.getElementById('inputId').value.trim();
        if (!id) { alert('Ingresa un ID de Vivienda'); return; }
        document.getElementById('msg').innerText = 'Generando...';
        google.script.run
          .withSuccessHandler(function(res) {
            document.getElementById('msg').innerText = '';
            if (res.error) { alert(res.error); return; }
            var div = document.createElement('div');
            div.innerHTML = res.html;
            document.body.innerHTML = '';
            document.body.appendChild(div);
          })
          .withFailureHandler(function(err) {
            document.getElementById('msg').innerText = 'Error: ' + err.message;
          })
          .generarEstadoCuentaHTML(id, document.getElementById('inputFechaInicio').value, document.getElementById('inputFechaFin').value);
      }
    </script>
  `;
  const output = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(output, "Estado de Cuenta");
}

function generarEstadoCuentaHTML(idBusqueda, strFechaInicio, strFechaFin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  idBusqueda = idBusqueda.toUpperCase().trim();
  const fechaInicio = strFechaInicio ? new Date(strFechaInicio + "T00:00:00") : null;
  const fechaFin = strFechaFin ? new Date(strFechaFin + "T23:59:59") : null;
  const fechaHoy = new Date();
  const diaHoy = fechaHoy.getDate();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreProp = "No encontrado";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][4].toString().toUpperCase().trim() === idBusqueda) {
      nombreProp = datosProp[i][0];
      break;
    }
  }

  const datosCargos = sheetCargos.getDataRange().getValues();
  let htmlCargos = "";
  let totalCargos = 0;
  let totalPendiente = 0;

  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1].toString().toUpperCase().trim() === idBusqueda) {
      const fechaCargo = datosCargos[i][0];
      if (fechaInicio && fechaCargo instanceof Date && fechaCargo < fechaInicio) continue;
      if (fechaFin && fechaCargo instanceof Date && fechaCargo > fechaFin) continue;

      const mes = datosCargos[i][5];
      const concepto = datosCargos[i][2];
      let monto = parseFloat(datosCargos[i][3]) || 0;
      const estatus = datosCargos[i][4];
      let montoStr = "$" + monto.toFixed(2);

      if (estatus === "Pendiente" && concepto === "Mantenimiento" && monto === 500) {
        const esMesAnterior = mes !== mesActualStr;
        if (esMesAnterior || diaHoy > 10) {
          montoStr = "$500 → <b style='color:#dc2626;'>$550</b> <span style='color:#92400e;font-size:11px;'>(+Recargo)</span>";
          monto = 550;
        }
      }

      const color = estatus === "Pendiente" ? "#dc2626" : "#16a34a";
      totalCargos += monto;
      if (estatus === "Pendiente") totalPendiente += monto;
      htmlCargos += `<tr><td>${mes}</td><td>${concepto}</td><td>${montoStr}</td><td style="color:${color};font-weight:bold;">${estatus}</td></tr>`;
    }
  }

  const datosHistorial = sheetHistorial.getDataRange().getValues();
  let htmlPagos = "";
  let totalPagadoHistorial = 0;
  const partes = idBusqueda.split(" ");
  const edificio = partes[0];
  const depto = partes.slice(1).join(" ");

  for (let i = 1; i < datosHistorial.length; i++) {
    const fecha = datosHistorial[i][0];
    if (fechaInicio && fecha instanceof Date && fecha < fechaInicio) continue;
    if (fechaFin && fecha instanceof Date && fecha > fechaFin) continue;

    const filaEdificio = (datosHistorial[i][1] || "").toString().toUpperCase().trim();
    const filaDepto = (datosHistorial[i][2] || "").toString().trim();
    if (filaEdificio === edificio && filaDepto === depto) {
      const monto = parseFloat(datosHistorial[i][6]) || 0;
      totalPagadoHistorial += monto;
      const fechaStr = fecha instanceof Date ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy") : fecha;
      htmlPagos += `<tr><td>${fechaStr}</td><td>${datosHistorial[i][4]}</td><td>${datosHistorial[i][5]}</td><td>$${monto.toFixed(2)}</td><td>${datosHistorial[i][7] || "---"}</td></tr>`;
    }
  }

  if (!htmlCargos && !htmlPagos) {
    return { error: "No se encontró información para " + idBusqueda + " en el rango seleccionado." };
  }

  const saldoActual = totalCargos - totalPagadoHistorial;

  let html = `
    <div style="font-family:sans-serif;padding:10px;">
      <h3 style="margin:0 0 5px 0;color:#1e3a8a;">${idBusqueda}</h3>
      <p style="margin:0 0 15px 0;color:#666;font-size:14px;">${nombreProp}</p>
      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <div style="flex:1;background:#f0fdf4;padding:10px;border-radius:8px;text-align:center;">
          <div style="font-size:11px;color:#666;">Total Cargado</div>
          <div style="font-size:18px;font-weight:bold;color:#16a34a;">$${totalCargos.toFixed(2)}</div>
        </div>
        <div style="flex:1;background:#fef2f2;padding:10px;border-radius:8px;text-align:center;">
          <div style="font-size:11px;color:#666;">Pendiente</div>
          <div style="font-size:18px;font-weight:bold;color:#dc2626;">$${totalPendiente.toFixed(2)}</div>
        </div>
        <div style="flex:1;background:#eff6ff;padding:10px;border-radius:8px;text-align:center;">
          <div style="font-size:11px;color:#666;">Pagado</div>
          <div style="font-size:18px;font-weight:bold;color:#2563eb;">$${totalPagadoHistorial.toFixed(2)}</div>
        </div>
      </div>
      <div style="text-align:center;padding:8px;border-radius:8px;font-weight:bold;margin-bottom:15px;${saldoActual <= 0 ? 'background:#f0fdf4;color:#16a34a;' : 'background:#fef2f2;color:#dc2626;'}">
        ${saldoActual <= 0 ? '✅ SALDO CUBIERTO' : '⚠️ SALDO PENDIENTE: $' + saldoActual.toFixed(2)}
      </div>`;

  if (htmlCargos) {
    html += `<h4 style="margin:10px 0 5px 0;color:#333;">Cargos</h4>
      <table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:15px;">
        <tr style="background:#1e3a8a;color:white;"><th style="padding:6px;">Mes</th><th style="padding:6px;">Concepto</th><th style="padding:6px;">Monto</th><th style="padding:6px;">Estatus</th></tr>
        ${htmlCargos}
      </table>`;
  }

  if (htmlPagos) {
    html += `<h4 style="margin:10px 0 5px 0;color:#333;">Historial de Pagos</h4>
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <tr style="background:#15803d;color:white;"><th style="padding:6px;">Fecha</th><th style="padding:6px;">Concepto</th><th style="padding:6px;">Detalle</th><th style="padding:6px;">Monto</th><th style="padding:6px;">Método</th></tr>
        ${htmlPagos}
      </table>`;
  }

  html += `</div>`;
  return { html: html };
}

function abrirFormularioPago() {
  const html = HtmlService.createHtmlOutputFromFile('FormPago')
    .setWidth(500)
    .setHeight(650)
    .setTitle('Registrar Pago');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function obtenerCargosPendientes(idVivienda) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const datos = sheetCargos.getDataRange().getValues();
  const id = idVivienda.toUpperCase().trim();
  const cargos = [];
  const fechaHoy = new Date();
  const diaHoy = fechaHoy.getDate();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1].toString().toUpperCase().trim() === id && datos[i][4] === "Pendiente") {
      const montoOriginal = parseFloat(datos[i][3]);
      let montoConRecargo = montoOriginal;
      if (datos[i][2] === "Mantenimiento" && montoOriginal === 500 && diaHoy > 10) {
        montoConRecargo = 550;
      }
      cargos.push({
        fila: i + 1,
        concepto: datos[i][2],
        mes: datos[i][5],
        montoOriginal: montoOriginal,
        montoConRecargo: montoConRecargo,
        monto: montoConRecargo
      });
    }
  }
  return cargos;
}
function procesarPagoDesdeHTML(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  const idVivienda = datos.idVivienda.toUpperCase().trim();
  const filaCargo = parseInt(datos.filaCargo);
  const montoPagado = parseFloat(datos.monto);
  const fechaPago = new Date(datos.fecha + "T12:00:00");
  const folio = datos.folio ? " (Folio: " + datos.folio + ")" : "";

  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreEncontrado = "No identificado";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][4].toString().toUpperCase().trim() === idVivienda) {
      nombreEncontrado = datosProp[i][0];
      break;
    }
  }

  // 1. Leer el cargo por número de fila
  const fila = sheetCargos.getRange(filaCargo, 1, 1, 6).getValues()[0];
  const concepto = fila[2];
  let montoDeudaActual = parseFloat(fila[3]) || 0;
  const mes = fila[5];
  if (!montoDeudaActual) throw new Error("Cargo no encontrado en la fila " + filaCargo);

  // 2. Lógica de Recargo Automático (solo para mantenimiento de $500)
  if (fechaPago.getDate() > 10 && concepto === "Mantenimiento" && montoDeudaActual === 500) {
    montoDeudaActual = 550;
  }

  // 3. Calcular Saldo
  const nuevoSaldo = montoDeudaActual - montoPagado;
  const estatusFinal = nuevoSaldo <= 0 ? "Pagado" : "Pendiente";
  const montoFinalCeldas = nuevoSaldo <= 0 ? montoDeudaActual : nuevoSaldo;

  // 4. Actualizar Hoja de Cargos
  sheetCargos.getRange(filaCargo, 4).setValue(montoFinalCeldas);
  sheetCargos.getRange(filaCargo, 5).setValue(estatusFinal);

  // 5. Registrar en Historial
  const partes = idVivienda.split(" ");
  const edificio = partes[0];
  const depto = partes.slice(1).join(" ");

  sheetHistorial.appendRow([
    fechaPago,
    edificio,
    depto,
    nombreEncontrado,
    concepto,
    "Pago de " + mes + folio,
    montoPagado,
    datos.metodo,
    datos.folio || "---"
  ]);

  return "¡Pago de " + nombreEncontrado + " registrado! Saldo: $" + (nuevoSaldo > 0 ? nuevoSaldo : 0);
}

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📋 Administración')
      .addItem('📊 Resumen Financiero', 'mostrarResumenFinanciero')
      .addSeparator()
      .addItem('➕ Agregar Cargo Individual', 'agregarCargoIndividual')
      .addItem('📅 Generar Cargos del Mes', 'generarCargosMes')
      .addSeparator()
      .addItem('📝 Registrar Pago (Formulario)', 'abrirFormularioPago')
      .addItem('🔍 Buscar Propietario', 'busquedaGlobalPropietario')
      .addItem('📄 Estado de Cuenta Completo', 'mostrarEstadoCuentaCompleto')
      .addItem('🏊 Lista de Bloqueo Piscina', 'generarListaBloqueo')
      .addToUi();
  } catch (e) {
    Logger.log("Error al crear menú: " + e.message);
  }
}

function mostrarResumenFinanciero() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const ui = SpreadsheetApp.getUi();

  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const fechaHoy = new Date();
  const mesActual = nombresMeses[fechaHoy.getMonth()];

  const datosCargos = sheetCargos.getDataRange().getValues();
  let totalPendiente = 0;
  let totalPagado = 0;
  let pendientesEsteMes = 0;
  let totalEsteMes = 0;

  for (let i = 1; i < datosCargos.length; i++) {
    const monto = parseFloat(datosCargos[i][3]) || 0;
    const estatus = datosCargos[i][4];
    const mes = datosCargos[i][5];

    if (estatus === "Pendiente") totalPendiente += monto;
    if (estatus === "Pagado") totalPagado += monto;
    if (mes === mesActual) {
      totalEsteMes += monto;
      if (estatus === "Pendiente") pendientesEsteMes++;
    }
  }

  const totalGeneral = totalPendiente + totalPagado;
  const cobranza = totalGeneral > 0 ? Math.round((totalPagado / totalGeneral) * 100) : 0;

  const datosHistorial = sheetHistorial.getDataRange().getValues();
  let cobradoEsteMes = 0;
  for (let i = 1; i < datosHistorial.length; i++) {
    const fecha = new Date(datosHistorial[i][0]);
    if (fecha.getMonth() === fechaHoy.getMonth() && fecha.getFullYear() === fechaHoy.getFullYear()) {
      cobradoEsteMes += parseFloat(datosHistorial[i][6]) || 0;
    }
  }

  const html = `
    <div style="font-family: sans-serif; padding: 20px;">
      <h3 style="margin:0 0 15px 0; color:#1e3a8a;">Resumen Financiero</h3>
      <p style="font-size:12px; color:#666; margin-bottom:20px;">${mesActual} ${fechaHoy.getFullYear()}</p>
      <table style="width:100%; border-collapse:collapse;">
        <tr><td style="padding:8px 0; font-weight:bold;">Total Histórico Cobrado</td><td style="text-align:right; color:#16a34a; font-weight:bold;">$${totalPagado.toLocaleString()}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Total Pendiente por Cobrar</td><td style="text-align:right; color:#dc2626; font-weight:bold;">$${totalPendiente.toLocaleString()}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Eficiencia de Cobranza</td><td style="text-align:right; font-weight:bold;">${cobranza}%</td></tr>
        <tr><td colspan="2"><hr style="border:0; border-top:1px solid #ddd;"></td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Cobrado en ${mesActual}</td><td style="text-align:right; color:#16a34a; font-weight:bold;">$${cobradoEsteMes.toLocaleString()}</td></tr>
        <tr><td style="padding:8px 0; font-weight:bold;">Unidades Pendientes en ${mesActual}</td><td style="text-align:right; color:#dc2626; font-weight:bold;">${pendientesEsteMes}</td></tr>
      </table>
    </div>
  `;

  const output = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(300);
  ui.showModalDialog(output, "📊 Dashboard");
}

function obtenerEstadoCuentaPorNombre(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const sheetCargos = ss.getSheetByName("Cargos");

  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  const nombreBusqueda = nombre.toUpperCase().trim();
  let idEncontrado = "";
  let nombreEncontrado = "";

  for (let i = 0; i < datosProp.length; i++) {
    const nombreProp = (datosProp[i][0] || "").toString().toUpperCase().trim();
    if (nombreProp.includes(nombreBusqueda)) {
      idEncontrado = datosProp[i][4];
      nombreEncontrado = datosProp[i][0];
      break;
    }
  }

  if (!idEncontrado) return { error: true };

  const fechaHoy = new Date();
  const diaActual = fechaHoy.getDate();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  const cargos = sheetCargos.getDataRange().getValues();
  let totalAdeudo = 0;
  let desglose = [];
  let contadorMesesPendientes = 0;

  for (let i = 1; i < cargos.length; i++) {
    if (cargos[i][1].toString().toUpperCase().trim() === idEncontrado.toUpperCase().trim() && cargos[i][4] === "Pendiente") {
      let montoCargo = parseFloat(cargos[i][3]);
      let mesCargo = cargos[i][5];
      let concepto = cargos[i][2];

      if (concepto === "Mantenimiento" && montoCargo === 500) {
        if (mesCargo !== mesActualStr || diaActual > 10) {
          montoCargo += 50;
        }
      }

      totalAdeudo += montoCargo;
      desglose.push({
        mes: mesCargo + (montoCargo > 500 ? " (+Recargo)" : ""),
        monto: montoCargo
      });
      contadorMesesPendientes++;
    }
  }

  return {
    nombre: nombreEncontrado,
    saldo: totalAdeudo,
    piscina: contadorMesesPendientes >= 2 ? "BLOQUEADO" : "ACCESO PERMITIDO",
    detalles: desglose
  };
}
