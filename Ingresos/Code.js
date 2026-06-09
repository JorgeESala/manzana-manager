/**
 * PROYECTO: AT-Version1
 * Automatización para MANZANA 80 ALDEA TULUM A.C.
 */

const FOLDER_ID = "1U2r8nB7DHfBIMnfzl366AI0j7VLE4jUP";

function obtenerBloqueActivo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const panel = ss.getSheetByName("Panel");
  if (!panel) return "MZ 17";
  const bloque = panel.getRange("B5").getValue();
  return bloque || "MZ 17";
}

function calcularRecargo(concepto, montoOriginal, mesCargo, fechaPago) {
  if (concepto !== "Mantenimiento" || montoOriginal !== 500) return 0;
  const meses = {"Enero":0,"Febrero":1,"Marzo":2,"Abril":3,"Mayo":4,"Junio":5,
                 "Julio":6,"Agosto":7,"Septiembre":8,"Octubre":9,"Noviembre":10,"Diciembre":11};
  const partes = (mesCargo || "").trim().split(/\s+/);
  if (partes.length !== 2) return 0;
  const mesNum = meses[partes[0]];
  const anio = parseInt(partes[1]);
  if (isNaN(mesNum) || isNaN(anio)) return 0;
  const limite = new Date(anio, mesNum, 10, 23, 59, 59);
  return fechaPago > limite ? 50 : 0;
}

// ==========================================
// 1. BOTÓN HOJA "RECIBOS": GUARDAR, PDF Y CORREO (GMAIL WEB)
// ==========================================
function generarRecibo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRecibos = ss.getSheetByName("Recibos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  const idVivienda = sheetRecibos.getRange("B3").getValue();
  const mesAPagar = sheetRecibos.getRange("G2").getValue();
  const montoPagado = sheetRecibos.getRange("F18").getValue();
  const formaPago = sheetRecibos.getRange("G5").getValue();
  const referencia = sheetRecibos.getRange("G6").getValue();

  if (!idVivienda || montoPagado <= 0) {
    SpreadsheetApp.getUi().alert("⚠️ Selecciona un departamento y un monto válido.");
    return;
  }

  // Read Propietarios columns D through H (Nombre, Tel, Correo, ID Vivienda, Manzana)
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreReal = "Nombre no encontrado";
  let bloqueProp = "MZ 17";

  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === idVivienda) {
      nombreReal = datosProp[i][0];
      bloqueProp = (datosProp[i][4] || "MZ 17").toString().trim();
      break;
    }
  }

  // Update Cargos status
  const datosCargos = sheetCargos.getDataRange().getValues();
  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1] === idVivienda && datosCargos[i][5] === mesAPagar && datosCargos[i][4] === "Pendiente") {
      sheetCargos.getRange(i + 1, 5).setValue("Pagado");
      sheetCargos.getRange(i + 1, 8).setValue(0);
      break;
    }
  }

  const fechaHoy = new Date();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];
  const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
  const esMesAnterior = mesAPagar !== mesAnioActual;
  const esPagoTardio = fechaHoy.getDate() > 10;
  let recargo = (esMesAnterior || esPagoTardio) ? 50 : 0;

  const partes = idVivienda.split(" ");
  const edificio = partes[0];
  const depto = partes[1] || "";

  sheetHistorial.appendRow([
    fechaHoy, edificio, depto, nombreReal,
    "Mantenimiento", "Pago de " + mesAPagar,
    montoPagado, formaPago, referencia, bloqueProp
  ]);

  if (recargo > 0) {
    sheetHistorial.appendRow([
      fechaHoy, edificio, depto, nombreReal,
      "RECARGO", "Recargo $50 " + mesAPagar,
      recargo, "Automático", "---", bloqueProp
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
  const registroDeudas = {};

  for (let i = 1; i < datosCargos.length; i++) {
    const idVivienda = datosCargos[i][1];
    const estatus = datosCargos[i][4];
    const bloque = (datosCargos[i][6] || "").toString().trim();

    if (estatus === "Pendiente") {
      if (!registroDeudas[idVivienda]) {
        registroDeudas[idVivienda] = { meses: [], total: 0, bloque: bloque };
      }
      registroDeudas[idVivienda].meses.push(datosCargos[i][5]);
      registroDeudas[idVivienda].total += parseFloat(datosCargos[i][3]);
    }
  }

  const listaBloqueo = [];
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();

  for (const id in registroDeudas) {
    if (registroDeudas[id].meses.length >= 2) {
      let nombre = "No encontrado";
      for (let j = 0; j < datosProp.length; j++) {
        if (datosProp[j][4] === id) {
          nombre = datosProp[j][0];
          break;
        }
      }
      listaBloqueo.push([
        id, nombre,
        registroDeudas[id].meses.length,
        "$" + registroDeudas[id].total,
        registroDeudas[id].meses.join(", "),
        registroDeudas[id].bloque
      ]);
    }
  }

  if (listaBloqueo.length === 0) {
    ui.alert("✅ ¡Excelente noticia!\n\nNo hay ninguna vivienda que deba 2 o más meses. Todos pueden usar la piscina.");
  } else {
    listaBloqueo.sort((a, b) => a[5].localeCompare(b[5]) || a[0].localeCompare(b[0]));

    let htmlContent = `<div style="font-family: sans-serif;">
      <p>Viviendas con <b>2 o más meses de adeudo</b>:</p>
      <table border="1" style="border-collapse: collapse; width: 100%; text-align: left;">
        <tr style="background-color: #f2f2f2;">
          <th style="padding: 8px;">Manzana</th>
          <th style="padding: 8px;">ID Vivienda</th>
          <th style="padding: 8px;">Propietario</th>
          <th style="padding: 8px;">Meses</th>
          <th style="padding: 8px;">Adeudo</th>
        </tr>`;

    listaBloqueo.forEach(row => {
      const bg = row[5] === "MZ 17" ? "#eff6ff" : "#f5f3ff";
      htmlContent += `<tr style="background-color:${bg};">
        <td style="padding: 8px; font-weight:bold;">${row[5]}</td>
        <td style="padding: 8px;">${row[0]}</td>
        <td style="padding: 8px;">${row[1]}</td>
        <td style="padding: 8px; color: red; font-weight: bold;">${row[2]}</td>
        <td style="padding: 8px;">${row[3]}</td>
      </tr>`;
    });

    htmlContent += `</table></div>`;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(450)
      .setTitle('BLOQUEO PISCINA');
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
  const sheetProp = ss.getSheetByName("Propietarios");

  // 1. Pedir el ID Vivienda (Edificio Depto)
  const deptoRes = ui.prompt("Nuevo Cargo Individual", "Ingrese el ID Vivienda (Ej: P2 101):", ui.ButtonSet.OK_CANCEL);
  if (deptoRes.getSelectedButton() != ui.Button.OK) return;
  const idVivienda = deptoRes.getResponseText().toUpperCase().trim();

  // Look up block from Propietarios
  const datosProp = sheetProp.getRange("D2:H" + sheetProp.getLastRow()).getValues();
  let bloqueEncontrado = "MZ 17";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === idVivienda) {
      bloqueEncontrado = (datosProp[i][4] || "MZ 17").toString().trim();
      break;
    }
  }

  // 2. Pedir el Motivo/Concepto
  const motivoRes = ui.prompt("Concepto", "Manzana detectada: " + bloqueEncontrado + "\n\nIngrese el motivo del cargo (ej. Multa, Reparación, Agua):", ui.ButtonSet.OK_CANCEL);
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

  sheetCargos.appendRow([
    new Date(),
    idVivienda,
    motivo,
    monto,
    "Pendiente",
    mes,
    bloqueEncontrado,
    monto
  ]);

  ss.toast("✅ Cargo de '" + motivo + "' registrado para " + idVivienda + " (" + bloqueEncontrado + ")");
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
  let saldoActual = 0;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === idVivienda && datos[i][5] === mes && datos[i][4] === "Pendiente") {
      filaEncontrada = i + 1;
      montoOriginal = datos[i][3];
      saldoActual = datos[i][7] || montoOriginal;
      break;
    }
  }

  if (filaEncontrada === -1) {
    ui.alert("No se encontró un cargo pendiente para " + idVivienda + " en " + mes);
    return;
  }

  // 4. Lógica de Recargo (si paga mes anterior o después del día 10)
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];
  const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
  const esMesAnterior = mes !== mesAnioActual;
  const esPagoTardio = fechaFinal.getDate() > 10;

  if (montoOriginal === 500 && (esMesAnterior || esPagoTardio) && saldoActual === montoOriginal) {
    saldoActual = 550;
    ss.toast("Se detectó pago extemporáneo. Total actualizado a $550", "Aviso");
  }

  // 5. Procesar el Pago
  const saldoRestante = saldoActual - montoPagado;
  const bloqueCargo = (datos[filaEncontrada - 1][6] || "MZ 17").toString().trim();

  if (saldoRestante <= 0) {
    sheetCargos.getRange(filaEncontrada, 5).setValue("Pagado");
    sheetCargos.getRange(filaEncontrada, 8).setValue(0);
    ui.alert("✅ Pago completo registrado.");
  } else {
    sheetCargos.getRange(filaEncontrada, 8).setValue(saldoRestante);
    ui.alert("⚠️ Pago parcial. El cargo queda pendiente por: $" + saldoRestante);
  }

  // 6. Registrar en Historial
  sheetHistorial.appendRow([fechaFinal, idVivienda.split(" ")[0], idVivienda.split(" ")[1], "---", "Mantenimiento", "Abono " + mes, montoPagado, "Transferencia/Efectivo", "---", bloqueCargo]);
}
function generarCargosMesPro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProp = ss.getSheetByName("Propietarios");
  const sheetCargos = ss.getSheetByName("Cargos");
  const ui = SpreadsheetApp.getUi();

  const mes = ui.prompt("Generar Cargos Mensuales", "Ingrese el mes (ej. Mayo):", ui.ButtonSet.OK).getResponseText();
  const monto = ui.prompt("Monto", "Monto de cuota:", ui.ButtonSet.OK).getResponseText();

  const datos = sheetProp.getRange("B2:H" + sheetProp.getLastRow()).getValues();
  const fecha = new Date();
  let countMZ17 = 0, countMZ19 = 0;

  datos.forEach(fila => {
    if (fila[0] !== "" && fila[1] !== "") {
      const edificio = fila[0];
      const depto = fila[1];
      const bloque = (fila[6] || "").toString().trim();

      const idVivienda = edificio + " " + depto;
      sheetCargos.appendRow([fecha, idVivienda, "Mantenimiento", monto, "Pendiente", mes, bloque, monto]);
      if (bloque === "MZ 17") countMZ17++;
      else if (bloque === "MZ 19") countMZ19++;
    }
  });

  ui.alert("✅ Cargos generados:\n\nMZ 17: " + countMZ17 + " unidades\nMZ 19: " + countMZ19 + " unidades");
}

function abrirFormularioCargoIndividual() {
  const html = HtmlService.createHtmlOutputFromFile('FormCargoIndividual')
    .setWidth(420)
    .setHeight(480)
    .setTitle('Agregar Cargo Individual');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function procesarCargoIndividual(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");

  const { manzana, idVivienda, concepto, monto, mes } = datos;

  sheetCargos.appendRow([
    new Date(),
    idVivienda,
    concepto,
    monto,
    "Pendiente",
    mes,
    manzana,
    monto
  ]);

  return "✅ Cargo de '" + concepto + "' ($" + monto + ") registrado para " + idVivienda + " (" + manzana + ")";
}

function abrirFormularioBusquedaGlobal() {
  const html = HtmlService.createHtmlOutputFromFile('FormBusquedaGlobal')
    .setWidth(420)
    .setHeight(300)
    .setTitle('Buscar Propietario');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function procesarBusquedaGlobal(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetPropietarios = ss.getSheetByName("Propietarios");

  const { manzana, idVivienda } = datos;
  const idBusqueda = idVivienda.toUpperCase().trim();
  const fechaHoy = new Date();
  const diaHoy = fechaHoy.getDate();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  let tablaHtml = "<table class='table table-bordered table-sm' style='font-size:14px;'>";
  tablaHtml += "<thead class='table-light'><tr><th>Mes</th><th>Concepto</th><th>Monto</th><th>Saldo</th><th>Estatus</th></tr></thead><tbody>";

  let encontrado = false;
  let deudaTotal = 0;

  const datosCargos = sheetCargos.getDataRange().getValues();
  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1].toString().toUpperCase() === idBusqueda) {
      const bloqueCargo = (datosCargos[i][6] || "").toString().trim();
      if (bloqueCargo !== manzana) continue;

      const mes = datosCargos[i][5];
      const concepto = datosCargos[i][2];
      const monto = parseFloat(datosCargos[i][3]) || 0;
      let saldo = parseFloat(datosCargos[i][7]) || monto;
      const estatus = datosCargos[i][4];

      let montoStr = "$" + monto;
      if (estatus === "Pendiente" && concepto === "Mantenimiento" && monto === 500) {
        const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
        const esMesAnterior = mes !== mesAnioActual;
        if (esMesAnterior || diaHoy > 10) {
          saldo = 550;
          montoStr = "<span style='text-decoration:line-through;color:#999;'>$500</span> → <b style='color:#dc2626;'>$550</b>";
        }
      }

      const colorEstatus = estatus === "Pendiente" ? "red" : "green";
      if (estatus === "Pendiente") deudaTotal += saldo;

      tablaHtml += "<tr><td>" + mes + "</td><td>" + concepto + "</td><td>" + montoStr + "</td><td>$" + saldo + "</td><td style='color:" + colorEstatus + "; font-weight:bold;'>" + estatus + "</td></tr>";
      encontrado = true;
    }
  }
  tablaHtml += "</tbody></table>";

  if (!encontrado) {
    return "<div class='alert alert-warning'>No se encontró información para: " + idBusqueda + " (" + manzana + ")</div>";
  }

  return "<h5>Resumen para: " + idBusqueda + " (" + manzana + ")</h5><p><b>Deuda Pendiente Total: $" + deudaTotal + "</b></p>" + tablaHtml;
}

function abrirFormularioCargos() {
  const html = HtmlService.createHtmlOutputFromFile('FormCargos')
    .setWidth(420)
    .setHeight(400)
    .setTitle('Generar Cargos Mensuales');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function verificarCargosExistentes(mesAnio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const datos = sheetCargos.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][5] === mesAnio) count++;
  }
  return { duplicados: count };
}

function generarCargosConFormulario(mesAnio, monto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProp = ss.getSheetByName("Propietarios");
  const sheetCargos = ss.getSheetByName("Cargos");

  const datos = sheetProp.getRange("B2:H" + sheetProp.getLastRow()).getValues();
  const fecha = new Date();
  let countMZ17 = 0, countMZ19 = 0;

  datos.forEach(fila => {
    if (fila[0] !== "" && fila[1] !== "") {
      const edificio = fila[0];
      const depto = fila[1];
      const bloque = (fila[6] || "").toString().trim();
      const idVivienda = edificio + " " + depto;
      sheetCargos.appendRow([fecha, idVivienda, "Mantenimiento", monto, "Pendiente", mesAnio, bloque, monto]);
      if (bloque === "MZ 17") countMZ17++;
      else if (bloque === "MZ 19") countMZ19++;
    }
  });

  return "✅ Cargos generados para " + mesAnio + ":\n\nMZ 17: " + countMZ17 + " unidades\nMZ 19: " + countMZ19 + " unidades";
}
function generarListaBloqueo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const datos = sheetCargos.getDataRange().getValues();
  const deudoresMZ17 = {};
  const deudoresMZ19 = {};

  for (let i = 1; i < datos.length; i++) {
    const idVivienda = datos[i][1].toString().toUpperCase().trim();
    const estatus = datos[i][4];
    const bloque = (datos[i][6] || "").toString().trim();

    if (estatus === "Pendiente") {
      if (bloque === "MZ 17") deudoresMZ17[idVivienda] = (deudoresMZ17[idVivienda] || 0) + 1;
      else if (bloque === "MZ 19") deudoresMZ19[idVivienda] = (deudoresMZ19[idVivienda] || 0) + 1;
    }
  }

  let listaMZ17 = [], listaMZ19 = [];
  for (const id in deudoresMZ17) { if (deudoresMZ17[id] >= 2) listaMZ17.push([id, deudoresMZ17[id]]); }
  for (const id in deudoresMZ19) { if (deudoresMZ19[id] >= 2) listaMZ19.push([id, deudoresMZ19[id]]); }

  if (listaMZ17.length === 0 && listaMZ19.length === 0) {
    SpreadsheetApp.getUi().alert("✅ ¡Excelente!\n\nNo hay ninguna vivienda que deba 2 o más meses. Todos pueden usar la piscina.");
    return;
  }

  let msj = "🏊 LISTA DE BLOQUEO PISCINA\n\n";
  if (listaMZ17.length > 0) {
    msj += "=== MZ 17 ===\n";
    listaMZ17.forEach(res => { msj += "  " + res[0] + " — " + res[1] + " meses pendientes\n"; });
    msj += "\n";
  }
  if (listaMZ19.length > 0) {
    msj += "=== MZ 19 ===\n";
    listaMZ19.forEach(res => { msj += "  " + res[0] + " — " + res[1] + " meses pendientes\n"; });
  }

  SpreadsheetApp.getUi().alert(msj);
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

  // 1. Buscar Nombre y Manzana desde Propietarios (D2:H incluye Manzana en col H)
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreEncontrado = "";
  let manzanaEncontrada = "MZ 17";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === idBusqueda) {
      nombreEncontrado = datosProp[i][0];
      manzanaEncontrada = (datosProp[i][4] || "MZ 17").toString().trim();
      break;
    }
  }
  if (!nombreEncontrado) return { error: true };

  // 2. Buscar Adeudos y calcular recargos "al vuelo" (sin filtro de bloque para Portal)
  const cargos = sheetCargos.getDataRange().getValues();
  let totalAdeudo = 0;
  let desglose = [];
  let contadorMesesPendientes = 0;

  for (let i = 1; i < cargos.length; i++) {
    if (cargos[i][1].toString().toUpperCase().trim() === idBusqueda && cargos[i][4] === "Pendiente") {
      const bloqueCargo = (cargos[i][6] || "").toString().trim();
      if (bloqueCargo !== manzanaEncontrada) continue;
      let montoCargo = parseFloat(cargos[i][3]);
      let mesCargo = cargos[i][5];
      let concepto = cargos[i][2];

      const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
      const recargo = calcularRecargo(concepto, montoCargo, mesCargo, fechaHoy);
      if (recargo > 0) {
        montoCargo += 50;
      }

      totalAdeudo += montoCargo;
      desglose.push({
        mes: mesCargo + (recargo > 0 ? " (+Recargo)" : ""),
        monto: montoCargo
      });
      contadorMesesPendientes++;
    }
  }

  return {
    nombre: nombreEncontrado,
    saldo: totalAdeudo,
    piscina: contadorMesesPendientes >= 2 ? "BLOQUEADO" : "ACCESO PERMITIDO",
    detalles: desglose,
    bloque: manzanaEncontrada
  };
}

function busquedaGlobalPropietario() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const respuesta = ui.prompt("Búsqueda Global", "Ingrese Edificio y Depto (Ej: P2 101):", ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() != ui.Button.OK) return;
  const idBusqueda = respuesta.getResponseText().toUpperCase();

  const fechaHoy = new Date();
  const diaHoy = fechaHoy.getDate();
  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const mesActualStr = nombresMeses[fechaHoy.getMonth()];

  const datosCargos = ss.getSheetByName("Cargos").getDataRange().getValues();
  let tablaHtml = "<table border='1' style='width:100%; border-collapse:collapse; font-family:Arial;'>";
  tablaHtml += "<tr style='background-color:#4a86e8; color:white;'><th>Mes</th><th>Concepto</th><th>Monto</th><th>Saldo</th><th>Estatus</th></tr>";

  let encontrado = false;
  let deudaTotal = 0;

  for (let i = 1; i < datosCargos.length; i++) {
    if (datosCargos[i][1].toString().toUpperCase() === idBusqueda) {
      const mes = datosCargos[i][5];
      const concepto = datosCargos[i][2];
      const monto = parseFloat(datosCargos[i][3]) || 0;
      let saldo = parseFloat(datosCargos[i][7]) || monto;
      const estatus = datosCargos[i][4];

      // Apply recargo display for pending maintenance charges
      let montoStr = "$" + monto;
      if (estatus === "Pendiente" && concepto === "Mantenimiento" && monto === 500) {
        const recargo = calcularRecargo(concepto, monto, mes, fechaHoy);
        if (recargo > 0) {
          saldo = 550;
          montoStr = "<span style='text-decoration:line-through;color:#999;'>$500</span> → <b style='color:#dc2626;'>$550</b>";
        }
      }

      const colorEstatus = estatus === "Pendiente" ? "red" : "green";
      if (estatus === "Pendiente") deudaTotal += saldo;

      tablaHtml += `<tr><td>${mes}</td><td>${concepto}</td><td>${montoStr}</td><td>$${saldo}</td><td style="color:${colorEstatus}; font-weight:bold;">${estatus}</td></tr>`;
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
    .setWidth(550)
    .setHeight(400);

  ui.showModalDialog(htmlOutput, "Estado de Cuenta Detallado");
}

function mostrarEstadoCuentaCompleto() {
  const html = `
    <div style="font-family:sans-serif;padding:20px;">
      <h3 style="margin:0 0 15px 0;color:#1e3a8a;">Estado de Cuenta por Fechas</h3>
      <div class="mb-3">
        <label style="display:block;font-size:12px;font-weight:bold;margin-bottom:4px;color:#333;">Manzana</label>
        <select id="inputManzana" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;">
          <option value="MZ 17">MZ 17</option>
          <option value="MZ 19">MZ 19</option>
        </select>
      </div>
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
          .generarEstadoCuentaHTML(id, document.getElementById('inputFechaInicio').value, document.getElementById('inputFechaFin').value, document.getElementById('inputManzana').value);
      }
    </script>
  `;
  const output = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(output, "Estado de Cuenta por Fechas");
}

function generarEstadoCuentaHTML(idBusqueda, strFechaInicio, strFechaFin, manzana) {
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
  let manzanaProp = null;
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === idBusqueda && datosProp[i][4] === manzana) {
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
      const bloqueCargo = (datosCargos[i][6] || "").toString().trim();
      if (bloqueCargo !== manzana) continue;
      const fechaCargo = datosCargos[i][0];
      if (fechaInicio && fechaCargo instanceof Date && fechaCargo < fechaInicio) continue;
      if (fechaFin && fechaCargo instanceof Date && fechaCargo > fechaFin) continue;

      const mes = datosCargos[i][5];
      const concepto = datosCargos[i][2];
      let monto = parseFloat(datosCargos[i][3]) || 0;
      const estatus = datosCargos[i][4];
      let montoStr = "$" + monto.toFixed(2);

      if (estatus === "Pendiente" && concepto === "Mantenimiento" && monto === 500) {
        const recargo = calcularRecargo(concepto, monto, mes, fechaHoy);
        if (recargo > 0) {
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
    const filaManzana = (datosHistorial[i][9] || "").toString().trim();
    if (filaEdificio === edificio && filaDepto === depto && filaManzana === manzana) {
      const monto = parseFloat(datosHistorial[i][6]) || 0;
      totalPagadoHistorial += monto;
      const fechaStr = fecha instanceof Date ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy") : fecha;
      htmlPagos += `<tr><td>${fechaStr}</td><td>${datosHistorial[i][4]}</td><td>${datosHistorial[i][5]}</td><td>$${monto.toFixed(2)}</td><td>${datosHistorial[i][7] || "---"}</td></tr>`;
    }
  }

  if (!htmlCargos && !htmlPagos) {
    return { error: "No se encontró información para " + idBusqueda + " (" + manzana + ") en el rango seleccionado." };
  }

  const saldoActual = totalCargos - totalPagadoHistorial;

  let html = `
    <div style="font-family:sans-serif;padding:10px;">
      <h3 style="margin:0 0 5px 0;color:#1e3a8a;">${idBusqueda} <span style="font-size:12px;color:#666;">(${manzana})</span></h3>
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
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const datos = sheetCargos.getDataRange().getValues();
  const id = idVivienda.toUpperCase().trim();
  const cargos = [];
  const fechaHoy = new Date();
  const diaHoy = fechaHoy.getDate();

  // Look up block from Propietarios
  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let bloqueEncontrado = "MZ 17";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === id) {
      bloqueEncontrado = (datosProp[i][4] || "MZ 17").toString().trim();
      break;
    }
  }

  for (let i = 1; i < datos.length; i++) {
    const bloqueCargo = (datos[i][6] || "").toString().trim();
    if (datos[i][1].toString().toUpperCase().trim() === id && datos[i][4] === "Pendiente" && bloqueCargo === bloqueEncontrado) {
      const montoOriginal = parseFloat(datos[i][3]);
      let saldoActual = parseFloat(datos[i][7]) || montoOriginal;
      const recargoHoy = calcularRecargo(datos[i][2], montoOriginal, datos[i][5], fechaHoy);
      if (recargoHoy > 0 && saldoActual === montoOriginal) {
        saldoActual = 550;
      }
      cargos.push({
        fila: i + 1,
        concepto: datos[i][2],
        mes: datos[i][5],
        montoOriginal: montoOriginal,
        tieneRecargo: recargoHoy > 0,
        saldoActual: saldoActual,
        monto: saldoActual
      });
    }
  }
  return { cargos: cargos, bloque: bloqueEncontrado };
}
function procesarPagoDesdeHTML(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCargos = ss.getSheetByName("Cargos");
  const sheetHistorial = ss.getSheetByName("Historial");
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const fechaHoy = new Date();

  const idVivienda = datos.idVivienda.toUpperCase().trim();
  const filaCargo = parseInt(datos.filaCargo);
  const montoPagado = parseFloat(datos.monto);
  const fechaPago = new Date(datos.fecha + "T12:00:00");
  const folio = datos.folio ? " (Folio: " + datos.folio + ")" : "";

  const datosProp = sheetPropietarios.getRange("D2:H" + sheetPropietarios.getLastRow()).getValues();
  let nombreEncontrado = "No identificado";
  let bloqueProp = "MZ 17";
  for (let i = 0; i < datosProp.length; i++) {
    if (datosProp[i][3].toString().toUpperCase().trim() === idVivienda) {
      nombreEncontrado = datosProp[i][0];
      bloqueProp = (datosProp[i][4] || "MZ 17").toString().trim();
      break;
    }
  }

  // 1. Leer el cargo por número de fila
  const fila = sheetCargos.getRange(filaCargo, 1, 1, 8).getValues()[0];
  const concepto = fila[2];
  const montoOriginal = parseFloat(fila[3]) || 0;
  let saldoActual = parseFloat(fila[7]) || montoOriginal;
  const mes = fila[5];
  const bloqueCargo = (fila[6] || bloqueProp).toString().trim();
  if (!montoOriginal) throw new Error("Cargo no encontrado en la fila " + filaCargo);

  // 2. Lógica de Recargo Automático (solo para mantenimiento de $500)
  const recargo = calcularRecargo(concepto, montoOriginal, mes, fechaPago);
  if (recargo > 0 && saldoActual === montoOriginal) {
    saldoActual = 550;
  }

  // 3. Calcular Saldo
  const nuevoSaldo = saldoActual - montoPagado;
  const estatusFinal = nuevoSaldo <= 0 ? "Pagado" : "Pendiente";

  // 4. Actualizar Hoja de Cargos — only modify Saldo (col 8), never modify Monto (col 4)
  sheetCargos.getRange(filaCargo, 8).setValue(Math.max(0, nuevoSaldo));
  sheetCargos.getRange(filaCargo, 5).setValue(estatusFinal);

  // 5. Registrar en Historial
  const partes = idVivienda.split(" ");
  const edificio = partes[0];
  const depto = partes.slice(1).join(" ");

  sheetHistorial.appendRow([
    fechaPago, edificio, depto, nombreEncontrado,
    concepto, "Pago de " + mes + folio,
    montoPagado, datos.metodo, datos.folio || "---", bloqueCargo
  ]);

  return "¡Pago de " + nombreEncontrado + " registrado! Saldo: $" + (nuevoSaldo > 0 ? nuevoSaldo : 0);
}

function abrirFormularioPropietario() {
  const html = HtmlService.createHtmlOutputFromFile('FormPropietario')
    .setWidth(450)
    .setHeight(580)
    .setTitle('Agregar Propietario');
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function procesarPropietarioHTML(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Propietarios");

  const edificio = datos.edificio.toUpperCase();
  const depto = datos.depto;
  const nombre = datos.nombre;
  const telefono = datos.telefono || "";
  const correo = datos.correo || "";
  const bloque = datos.bloque;
  const idVivienda = edificio + " " + depto;

  // Check for duplicate: same idVivienda + same block
  const datosExistentes = sheet.getRange("H2:I" + sheet.getLastRow()).getValues();
  for (let i = 0; i < datosExistentes.length; i++) {
    const idExistente = (datosExistentes[i][0] || "").toString().toUpperCase().trim();
    const bloqueExistente = (datosExistentes[i][1] || "").toString().trim();
    if (idExistente === idVivienda.toUpperCase().trim() && bloqueExistente === bloque) {
      throw new Error("Ya existe un propietario con ID " + idVivienda + " en bloque " + bloque + ".");
    }
  }

  // Generate next ID
  const lastRow = sheet.getLastRow();
  let nextId = 1;
  if (lastRow > 1) {
    const ids = sheet.getRange("A2:A" + lastRow).getValues();
    for (let i = 0; i < ids.length; i++) {
      const num = parseInt(ids[i][0]);
      if (!isNaN(num) && num >= nextId) nextId = num + 1;
    }
  }

  sheet.appendRow([nextId, edificio, depto, nombre, telefono, correo, "Activo", idVivienda, bloque]);

  return "✅ Propietario '" + nombre + "' agregado en " + idVivienda + " (" + bloque + ")";
}

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📋 Administración')
      .addItem('📊 Resumen Financiero', 'mostrarResumenFinanciero')
      .addSeparator()
      .addItem('➕ Agregar Propietario', 'abrirFormularioPropietario')
      .addItem('➕ Agregar Cargo Individual', 'abrirFormularioCargoIndividual')
      .addItem('📅 Generar Cargos de Mantenimiento', 'abrirFormularioCargos')
      .addSeparator()
      .addItem('📝 Registrar Pago (Formulario)', 'abrirFormularioPago')
      .addItem('🔍 Buscar Propietario', 'abrirFormularioBusquedaGlobal')
      .addItem('📄 Estado de Cuenta por Fechas', 'mostrarEstadoCuentaCompleto')
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
  let totalCargadoMZ17 = 0, totalCargadoMZ19 = 0;
  let saldoPendienteMZ17 = 0, saldoPendienteMZ19 = 0;
  let pendientesEsteMesMZ17 = 0, pendientesEsteMesMZ19 = 0;
  let recargosPendientesMZ17 = 0, recargosPendientesMZ19 = 0;

  const diaActual = fechaHoy.getDate();
  const mesAnioActual = mesActual + " " + fechaHoy.getFullYear();

  for (let i = 1; i < datosCargos.length; i++) {
    const bloque = (datosCargos[i][6] || "").toString().trim();
    const monto = parseFloat(datosCargos[i][3]) || 0;
    const saldo = parseFloat(datosCargos[i][7]) || 0;
    const estatus = datosCargos[i][4];
    const mes = datosCargos[i][5];
    const concepto = datosCargos[i][2];

    let recargoAplicable = 0;
    if (estatus === "Pendiente" && concepto === "Mantenimiento" && monto === 500 && saldo === monto) {
      recargoAplicable = calcularRecargo(concepto, monto, mes, fechaHoy);
    }

    if (bloque === "MZ 17") {
      totalCargadoMZ17 += monto;
      if (estatus === "Pendiente") saldoPendienteMZ17 += saldo;
      if (estatus === "Pendiente") recargosPendientesMZ17 += recargoAplicable;
      if (mes === mesAnioActual && estatus === "Pendiente") pendientesEsteMesMZ17++;
    } else if (bloque === "MZ 19") {
      totalCargadoMZ19 += monto;
      if (estatus === "Pendiente") saldoPendienteMZ19 += saldo;
      if (estatus === "Pendiente") recargosPendientesMZ19 += recargoAplicable;
      if (mes === mesAnioActual && estatus === "Pendiente") pendientesEsteMesMZ19++;
    }
  }

  // Total paid from Historial
  let totalPagadoMZ17 = 0, totalPagadoMZ19 = 0;
  const datosHistorial = sheetHistorial.getDataRange().getValues();
  let cobradoEsteMesMZ17 = 0, cobradoEsteMesMZ19 = 0;

  for (let i = 1; i < datosHistorial.length; i++) {
    const bloque = (datosHistorial[i][9] || "").toString().trim();
    const monto = parseFloat(datosHistorial[i][6]) || 0;
    const concepto = (datosHistorial[i][4] || "").toString().trim();
    const detalle = (datosHistorial[i][5] || "").toString();
    const esPagoDelMes = detalle.includes(mesAnioActual) && concepto === "Mantenimiento";

    if (bloque === "MZ 17") {
      totalPagadoMZ17 += monto;
      if (esPagoDelMes) cobradoEsteMesMZ17 += monto;
    } else if (bloque === "MZ 19") {
      totalPagadoMZ19 += monto;
      if (esPagoDelMes) cobradoEsteMesMZ19 += monto;
    }
  }

  const cobranzaMZ17 = totalCargadoMZ17 > 0 ? Math.round((totalPagadoMZ17 / totalCargadoMZ17) * 100) : 0;
  const cobranzaMZ19 = totalCargadoMZ19 > 0 ? Math.round((totalPagadoMZ19 / totalCargadoMZ19) * 100) : 0;
  

  const html = `
    <div style="font-family: sans-serif; padding: 20px;">
      <h3 style="margin:0 0 15px 0; color:#1e3a8a;">Resumen Financiero</h3>
      <p style="font-size:12px; color:#666; margin-bottom:20px;">${mesActual} ${fechaHoy.getFullYear()}</p>
      <div style="display:flex;gap:10px;margin-bottom:15px;">
        <div style="flex:1;background:#eff6ff;padding:12px;border-radius:8px;">
          <div style="font-weight:bold;color:#1e3a8a;margin-bottom:8px;">MZ 17</div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Total Cargado:</b> <span>$${totalCargadoMZ17.toLocaleString()}</span></div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Total Pagado:</b> <span style="color:#16a34a;">$${totalPagadoMZ17.toLocaleString()}</span></div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Saldo Pendiente:</b> <span style="color:#dc2626;">$${saldoPendienteMZ17.toLocaleString()}</span></div>
          ${recargosPendientesMZ17 > 0 ? `<div style="font-size:13px;margin-bottom:4px;"><b>Recargos Pendientes:</b> <span style="color:#92400e;">+ $${recargosPendientesMZ17.toLocaleString()}</span></div>` : ''}
          <div style="font-size:13px;margin-bottom:4px;"><b>Cobranza:</b> ${cobranzaMZ17}%</div>
          <hr style="border:0;border-top:1px solid #ccc;margin:6px 0;">
          <div style="font-size:13px;margin-bottom:4px;"><b>Cobrado ${mesActual}:</b> <span style="color:#16a34a;">$${cobradoEsteMesMZ17.toLocaleString()}</span></div>
          <div style="font-size:13px;"><b>Pendientes este mes:</b> <span style="color:#dc2626;">${pendientesEsteMesMZ17}</span></div>
        </div>
        <div style="flex:1;background:#f5f3ff;padding:12px;border-radius:8px;">
          <div style="font-weight:bold;color:#674ea7;margin-bottom:8px;">MZ 19</div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Total Cargado:</b> <span>$${totalCargadoMZ19.toLocaleString()}</span></div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Total Pagado:</b> <span style="color:#16a34a;">$${totalPagadoMZ19.toLocaleString()}</span></div>
          <div style="font-size:13px;margin-bottom:4px;"><b>Saldo Pendiente:</b> <span style="color:#dc2626;">$${saldoPendienteMZ19.toLocaleString()}</span></div>
          ${recargosPendientesMZ19 > 0 ? `<div style="font-size:13px;margin-bottom:4px;"><b>Recargos Pendientes:</b> <span style="color:#92400e;">+ $${recargosPendientesMZ19.toLocaleString()}</span></div>` : ''}
          <div style="font-size:13px;margin-bottom:4px;"><b>Cobranza:</b> ${cobranzaMZ19}%</div>
          <hr style="border:0;border-top:1px solid #ccc;margin:6px 0;">
          <div style="font-size:13px;margin-bottom:4px;"><b>Cobrado ${mesActual}:</b> <span style="color:#16a34a;">$${cobradoEsteMesMZ19.toLocaleString()}</span></div>
          <div style="font-size:13px;"><b>Pendientes este mes:</b> <span style="color:#dc2626;">${pendientesEsteMesMZ19}</span></div>
        </div>
      </div>
    </div>
  `;

  const output = HtmlService.createHtmlOutput(html).setWidth(550).setHeight(350);
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
  let manzanaEncontrada = "MZ 17";

  for (let i = 0; i < datosProp.length; i++) {
    const nombreProp = (datosProp[i][0] || "").toString().toUpperCase().trim();
    if (nombreProp.includes(nombreBusqueda)) {
      idEncontrado = datosProp[i][3];
      nombreEncontrado = datosProp[i][0];
      manzanaEncontrada = (datosProp[i][4] || "MZ 17").toString().trim();
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
      const bloqueCargo = (cargos[i][6] || "").toString().trim();
      if (bloqueCargo !== manzanaEncontrada) continue;
      let montoCargo = parseFloat(cargos[i][3]);
      let mesCargo = cargos[i][5];
      let concepto = cargos[i][2];

      const mesAnioActual = mesActualStr + " " + fechaHoy.getFullYear();
      if (concepto === "Mantenimiento" && montoCargo === 500) {
        if (mesCargo !== mesAnioActual || diaActual > 10) {
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
    detalles: desglose,
    bloque: manzanaEncontrada
  };
}

function obtenerCoeficientes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPropietarios = ss.getSheetByName("Propietarios");
  const datosProp = sheetPropietarios.getRange("I2:I" + sheetPropietarios.getLastRow()).getValues();
  let countMZ17 = 0, countMZ19 = 0;
  for (let i = 0; i < datosProp.length; i++) {
    const b = (datosProp[i][0] || "").toString().trim();
    if (b === "MZ 17") countMZ17++;
    else if (b === "MZ 19") countMZ19++;
  }
  const total = countMZ17 + countMZ19;
  return {
    "MZ 17": { count: countMZ17, percent: total > 0 ? countMZ17 / total : 0.5 },
    "MZ 19": { count: countMZ19, percent: total > 0 ? countMZ19 / total : 0.5 },
    total: total
  };
}
