/**
 * ==============================================================================
 * VELORA BEAUTY - GOOGLE APPS SCRIPT BACKEND & META CONVERSIONS API (CAPI)
 * Hoja de Cálculo ID: 1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8
 * ==============================================================================
 * 
 * INSTRUCCIONES DE INSTALACIÓN Y CONFIGURACIÓN:
 * 1. Abre tu hoja de Google Sheets:
 *    https://docs.google.com/spreadsheets/d/1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8/edit
 * 2. En el menú superior: Extensiones > Apps Script.
 * 3. Reemplaza todo el contenido por este archivo y guarda (Ctrl + S).
 * 4. Configura tus credenciales seguras (Propiedades del script):
 *    - En el menú izquierdo de Apps Script, haz clic en el engranaje ⚙️ (Configuración del proyecto).
 *    - Baja hasta "Propiedades de la secuencia de comandos" y añade:
 *        * META_CAPI_ACCESS_TOKEN : Tu token permanente de Administrador de Eventos de Meta.
 *        * META_PIXEL_ID          : 1635217208321987 (o tu nuevo ID de Pixel/Conjunto de datos).
 *        * META_TEST_EVENT_CODE   : (Opcional) Código de prueba para depurar en Eventos de prueba (ej. TEST12345).
 *        * ADMIN_SECRET           : (Opcional) Clave maestra para autorizar cambios sensibles desde /admin.
 * 5. Haz clic en: Implementar > Nueva implementación.
 *    - Tipo: "Aplicación web".
 *    - Ejecutar como: "Yo" (tu cuenta de Google).
 *    - Quién tiene acceso: "Cualquier persona" (Anyone).
 * 6. Copia la URL generada (/exec) y pégala en la configuración de admin.html.
 */

var SPREADSHEET_ID = "1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8";
var DEFAULT_PIXEL_ID = "1635217208321987";
var SHEET_ORDERS = "Pedidos";
var SHEET_METRICS = "Métricas";

/**
 * FUNCIÓN DE CONFIGURACIÓN DIRECTA EN APPS SCRIPT:
 * Ejecuta esta función desde el editor de Google Apps Script para guardar tu token:
 * 1. Selecciona 'configurarVariableTokem' en el menú de funciones arriba.
 * 2. Haz clic en 'Ejecutar'.
 */
function configurarVariableTokem() {
  PropertiesService.getScriptProperties().setProperty(
    "tokem",
    "EAAdSX7uwHFUBSRusZCXXUxYqHxVebWuaZCsmPZCQUajy9sVEzTDcbLwo2vMKRZCKVobY9ujZAAZByaNyb9mCHZCV0mxgfaHySyvSHyZCLwOw2uaYpvizsI4WwaxbNjfbtiaqbtqNykHNbHK9sALUsTW5yQiu31dqFdS7LGeLfzZBJtvI2ulBZAg4OcbwPZA24Dr4Sqy5gZDZD"
  );
  PropertiesService.getScriptProperties().setProperty("META_PIXEL_ID", "1635217208321987");
  Logger.log("✅ Variable 'tokem' y 'META_PIXEL_ID' configuradas exitosamente en el servidor de Google Apps Script.");
}

// Catálogo oficial de precios para validación estricta en servidor
var PRICE_CATALOG = {
  1: { units: 1, price: 99.90, name: "1x VELORA 5 en 1" },
  2: { units: 2, price: 169.90, name: "2x VELORA 5 en 1 (Ahorro S/29.90)" },
  3: { units: 3, price: 229.90, name: "3x VELORA 5 en 1 (Ahorro S/69.80)" }
};

// Encabezados oficiales de la Hoja Pedidos (17 columnas estructuradas)
var HEADERS = [
  "ID Pedido",          // Col A (1)
  "Fecha y Hora",       // Col B (2)
  "Cliente",            // Col C (3)
  "WhatsApp",           // Col D (4)
  "DNI",                // Col E (5)
  "Departamento",       // Col F (6)
  "Agencia Shalom",     // Col G (7)
  "Producto",           // Col H (8)
  "Unidades",           // Col I (9)
  "Total (S/)",         // Col J (10)
  "Método Pago",        // Col K (11)
  "Estado Logístico",   // Col L (12) - Pendiente, Confirmado, En Camino, Entregado, Cancelado
  "Estado Pago",        // Col M (13) - Pendiente, Pagado
  "Fecha Pago",         // Col N (14) - Fecha real del ingreso Yape/Plin
  "Estado CAPI",        // Col O (15) - none, sent, failed
  "ID Evento CAPI",     // Col P (16) - purchase_VL-XXXX
  "Atribución (JSON)"   // Col Q (17) - fbp, fbc, UTMs, client IP, client UA
];

function getTargetSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Inicializar o actualizar automáticamente hojas y encabezados sin perder datos previos
 */
function setupSheet() {
  var ss = getTargetSpreadsheet();
  
  // 1. Configurar Hoja "Pedidos"
  var ordersSheet = ss.getSheetByName(SHEET_ORDERS);
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet(SHEET_ORDERS, 0);
  }
  
  var lastRow = ordersSheet.getLastRow();
  var lastCol = ordersSheet.getLastColumn();

  if (lastRow === 0) {
    ordersSheet.appendRow(HEADERS);
    var headerRange = ordersSheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0F172A");
    headerRange.setFontColor("#FFFFFF");
    headerRange.setFontFamily("Plus Jakarta Sans");
    headerRange.setHorizontalAlignment("center");
    ordersSheet.setFrozenRows(1);
    
    for (var i = 1; i <= HEADERS.length; i++) {
      ordersSheet.setColumnWidth(i, 160);
    }
    ordersSheet.setColumnWidth(3, 220); // Cliente
    ordersSheet.setColumnWidth(7, 300); // Agencia Shalom
  } else if (lastCol < HEADERS.length) {
    // Si la hoja ya existía con 12 columnas, añadir las nuevas columnas sin alterar los pedidos existentes
    var missingHeaders = HEADERS.slice(lastCol);
    var newHeaderRange = ordersSheet.getRange(1, lastCol + 1, 1, missingHeaders.length);
    newHeaderRange.setValues([missingHeaders]);
    newHeaderRange.setFontWeight("bold");
    newHeaderRange.setBackground("#0F172A");
    newHeaderRange.setFontColor("#FFFFFF");
    newHeaderRange.setFontFamily("Plus Jakarta Sans");
    newHeaderRange.setHorizontalAlignment("center");
    
    for (var j = lastCol + 1; j <= HEADERS.length; j++) {
      ordersSheet.setColumnWidth(j, 170);
    }
  }

  // 2. Configurar Hoja "Métricas"
  var metricsSheet = ss.getSheetByName(SHEET_METRICS);
  if (!metricsSheet) {
    metricsSheet = ss.insertSheet(SHEET_METRICS, 1);
    
    metricsSheet.getRange("A1:B1").merge().setValue("📊 PANEL DE CONTROL Y MÉTRICAS VELORA")
      .setFontWeight("bold").setFontSize(14).setBackground("#0B0F19").setFontColor("#22C55E").setHorizontalAlignment("center");
    
    var metricsData = [
      ["Métrica", "Fórmula Automática"],
      ["Ventas Registradas (S/)", '=IFERROR(SUM(Pedidos!J2:J), 0)'],
      ["Total Cobrado / Pagado (S/)", '=IFERROR(SUMIFS(Pedidos!J2:J, Pedidos!M2:M, "Pagado"), 0)'],
      ["Total de Pedidos Registrados", '=MAX(0, COUNTA(Pedidos!A2:A)-1)'],
      ["Pedidos con Pago Pendiente", '=COUNTIF(Pedidos!M2:M, "Pendiente")'],
      ["Pedidos Pagados Confirmados", '=COUNTIF(Pedidos!M2:M, "Pagado")'],
      ["Eventos CAPI Enviados", '=COUNTIF(Pedidos!O2:O, "sent")'],
      ["Pedidos en Logística Pendientes", '=COUNTIF(Pedidos!L2:L, "Pendiente")'],
      ["Pedidos Confirmados", '=COUNTIF(Pedidos!L2:L, "Confirmado")'],
      ["Pedidos En Camino", '=COUNTIF(Pedidos!L2:L, "En Camino")'],
      ["Pedidos Entregados", '=COUNTIF(Pedidos!L2:L, "Entregado")'],
      ["Pedidos Cancelados", '=COUNTIF(Pedidos!L2:L, "Cancelado")'],
      ["Ticket Promedio (S/)", '=IFERROR(AVERAGE(Pedidos!J2:J), 0)']
    ];
    
    metricsSheet.getRange(2, 1, metricsData.length, 2).setValues(metricsData);
    metricsSheet.getRange("A2:B2").setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
    metricsSheet.getRange("B3:B14").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right");
    metricsSheet.setColumnWidth(1, 280);
    metricsSheet.setColumnWidth(2, 220);
  }
  
  return { status: "success", message: "Hojas y encabezados configurados correctamente" };
}

// ==========================================================
// HELPERS DE NORMALIZACIÓN Y SEGURIDAD PARA META CAPI
// ==========================================================

function stripAccents(str) {
  if (!str) return "";
  return String(str)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/ñ/gi, "n")
    .toLowerCase()
    .trim();
}

function hashSHA256(text) {
  if (!text) return "";
  var clean = String(text).trim().toLowerCase();
  if (!clean) return "";
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clean, Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < raw.length; i++) {
    var b = raw[i];
    if (b < 0) b += 256;
    var h = b.toString(16);
    if (h.length === 1) h = "0" + h;
    hex += h;
  }
  return hex;
}

/**
 * Enviar evento Purchase a Meta Conversions API (Graph API v21.0)
 */
function sendCapiPurchase(order, paidTimestamp) {
  var props = PropertiesService.getScriptProperties();
  var accessToken = props.getProperty("tokem") || props.getProperty("META_CAPI_ACCESS_TOKEN");
  var pixelId = props.getProperty("META_PIXEL_ID") || DEFAULT_PIXEL_ID;
  var testEventCode = props.getProperty("META_TEST_EVENT_CODE");

  if (!accessToken || accessToken.trim() === "") {
    return {
      status: "failed",
      event_id: "purchase_" + String(order.id).trim(),
      message: "Variable 'tokem' (Access Token de Meta CAPI) no configurada en Propiedades del Script de Google Apps Script."
    };
  }

  // 1. Normalización de datos del usuario (SHA-256)
  var userData = {};

  // Teléfono a formato E.164 (+51 Perú)
  var rawPhone = String(order.phone || "").replace(/\D/g, "");
  if (rawPhone.length === 9 && rawPhone.charAt(0) === '9') {
    rawPhone = "51" + rawPhone;
  }
  if (rawPhone) {
    userData.ph = [hashSHA256(rawPhone)];
  }

  // Nombre y Apellido (sin tildes, minúsculas)
  var cleanFullName = stripAccents(order.name || "");
  var nameParts = cleanFullName.split(/\s+/).filter(function(p) { return p.length > 0; });
  if (nameParts.length > 0) {
    userData.fn = [hashSHA256(nameParts[0])];
    if (nameParts.length > 1) {
      var lastName = nameParts.slice(1).join("");
      userData.ln = [hashSHA256(lastName)];
    }
  }

  // Ciudad/Departamento (sin tildes, minúsculas)
  var cleanCity = stripAccents(order.city || "").replace(/[^a-z0-9]/g, "");
  if (cleanCity) {
    userData.ct = [hashSHA256(cleanCity)];
  }

  // País: Perú ('pe')
  userData.country = [hashSHA256("pe")];

  // Atribución del comprador original (NO del administrador)
  var attr = order.attribution || {};
  if (typeof attr === "string") {
    try { attr = JSON.parse(attr); } catch(e) { attr = {}; }
  }

  if (attr.fbp) userData.fbp = String(attr.fbp).trim();
  if (attr.fbc) userData.fbc = String(attr.fbc).trim();
  if (attr.client_user_agent) userData.client_user_agent = String(attr.client_user_agent).trim();
  if (attr.client_ip) userData.client_ip_address = String(attr.client_ip).trim();

  // NOTA CRÍTICA DE SEGURIDAD: NUNCA se envía DNI a Meta ni se usan datos del administrador

  var eventId = "purchase_" + String(order.id).trim();
  var eventEpochSeconds = paidTimestamp 
    ? Math.floor(new Date(paidTimestamp).getTime() / 1000) 
    : Math.floor(Date.now() / 1000);

  var eventPayload = {
    event_name: "Purchase",
    event_time: eventEpochSeconds,
    event_id: eventId,
    event_source_url: attr.landing_url || "https://velorabeautype.store/",
    action_source: "chat", // Venta concretada y verificada mediante WhatsApp Chat
    user_data: userData,
    custom_data: {
      currency: "PEN",
      value: Number(order.price) || 99.90,
      content_name: order.product || "Secadora 5 en 1 VELORA",
      content_type: "product",
      num_items: Number(order.units) || 1,
      content_ids: ["velora-5en1-" + (order.units || 1) + "u"]
    }
  };

  var postData = {
    data: [eventPayload]
  };

  if (testEventCode && testEventCode.trim() !== "") {
    postData.test_event_code = testEventCode.trim();
  }

  var url = "https://graph.facebook.com/v21.0/" + pixelId + "/events?access_token=" + encodeURIComponent(accessToken.trim());

  try {
    var response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(postData),
      muteHttpExceptions: true
    });

    var resCode = response.getResponseCode();
    var resText = response.getContentText();
    var resJson = {};
    try { resJson = JSON.parse(resText); } catch(errParse) {}

    if (resCode >= 200 && resCode < 300 && resJson.events_received > 0) {
      return {
        status: "sent",
        event_id: eventId,
        message: "Evento Purchase recibido exitosamente por Meta CAPI (events_received: " + resJson.events_received + ")",
        response: resJson
      };
    } else {
      return {
        status: "failed",
        event_id: eventId,
        message: "Meta Graph API error (HTTP " + resCode + "): " + (resJson.error ? resJson.error.message : resText),
        response: resJson
      };
    }
  } catch (errReq) {
    return {
      status: "failed",
      event_id: eventId,
      message: "Excepción al conectar con Meta CAPI: " + errReq.toString()
    };
  }
}

/**
 * Endpoint principal POST
 */
function doPost(e) {
  try {
    setupSheet();
    var ss = getTargetSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ORDERS);
    
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    var action = data.action || "create_order";

    return handleRequest(action, data, sheet);

  } catch (error) {
    return respondJSON({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * Endpoint principal GET
 */
function doGet(e) {
  try {
    setupSheet();
    var ss = getTargetSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ORDERS);
    
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || "get_orders";

    return handleRequest(action, params, sheet);

  } catch (error) {
    return respondJSON({
      status: "error",
      message: error.toString(),
      orders: []
    });
  }
}

/**
 * Enrutador de acciones
 */
function handleRequest(action, data, sheet) {

  // 1. LEER PEDIDOS
  if (action === "get_orders") {
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return respondJSON({ status: "success", count: 0, orders: [] });
    }
    
    var dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    var rawValues = dataRange.getValues();
    
    var orders = [];
    for (var i = rawValues.length - 1; i >= 0; i--) {
      var row = rawValues[i];
      if (!row[0]) continue;
      
      var attrObj = {};
      try {
        if (row[16]) attrObj = JSON.parse(row[16]);
      } catch(e) {}

      orders.push({
        id: String(row[0]),
        date: String(row[1]),
        name: String(row[2]),
        phone: String(row[3]).replace(/^'/, ''),
        dni: String(row[4]).replace(/^'/, ''),
        city: String(row[5]),
        address: String(row[6]),
        product: String(row[7]),
        units: Number(row[8]) || 1,
        price: Number(row[9]) || 0,
        payment: String(row[10]),
        status: String(row[11]) || "Pendiente",          // Estado logístico
        payment_status: String(row[12]) || "Pendiente",  // Estado pago
        paid_at: String(row[13] || ""),
        capi_status: String(row[14] || "none"),
        capi_event_id: String(row[15] || ""),
        attribution: attrObj
      });
    }

    return respondJSON({
      status: "success",
      count: orders.length,
      orders: orders
    });
  }

  // 2. ACTUALIZAR ESTADO LOGÍSTICO (Confirmado, En Camino, Entregado, etc.) - NUNCA ENVÍA PURCHASE
  if (action === "update_status") {
    var targetId = data.id;
    var newStatus = data.status;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var r = 0; r < idValues.length; r++) {
        if (String(idValues[r][0]).trim() === String(targetId).trim()) {
          sheet.getRange(r + 2, 12).setValue(newStatus); // Columna L = Estado Logístico
          return respondJSON({
            status: "success",
            message: "Estado logístico actualizado a " + newStatus,
            id: targetId,
            newStatus: newStatus
          });
        }
      }
    }
    return respondJSON({ status: "error", message: "Pedido no encontrado" });
  }

  // 3. MARCAR COMO PAGADO Y ENVIAR CAPI PURCHASE
  if (action === "mark_paid") {
    var orderIdToPay = data.id;
    var lastR = sheet.getLastRow();
    if (lastR <= 1) {
      return respondJSON({ status: "error", message: "No hay pedidos en la hoja" });
    }

    var rows = sheet.getRange(2, 1, lastR - 1, HEADERS.length).getValues();
    for (var k = 0; k < rows.length; k++) {
      var rData = rows[k];
      if (String(rData[0]).trim() === String(orderIdToPay).trim()) {
        var rowIndex = k + 2;
        var currentPaymentStatus = String(rData[12] || "Pendiente").trim();
        var currentCapiStatus = String(rData[14] || "none").trim();

        // Si ya estaba pagado, evitar compra duplicada bajo cualquier circunstancia
        if (currentPaymentStatus === "Pagado") {
          return respondJSON({
            status: "success",
            message: "El pedido #" + orderIdToPay + " ya fue marcado como Pagado anteriormente. No se envía ninguna compra adicional.",
            id: orderIdToPay,
            payment_status: "Pagado",
            capi_status: currentCapiStatus,
            capi_event_id: String(rData[15] || "")
          });
        }

        var paidDateStr = Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");
        sheet.getRange(rowIndex, 13).setValue("Pagado");      // Col M = Estado Pago
        sheet.getRange(rowIndex, 14).setValue(paidDateStr);   // Col N = Fecha Pago

        var orderObj = {
          id: String(rData[0]),
          date: String(rData[1]),
          name: String(rData[2]),
          phone: String(rData[3]).replace(/^'/, ''),
          dni: String(rData[4]).replace(/^'/, ''),
          city: String(rData[5]),
          address: String(rData[6]),
          product: String(rData[7]),
          units: Number(rData[8]) || 1,
          price: Number(rData[9]) || 99.90,
          payment: String(rData[10]),
          attribution: rData[16] || {}
        };

        // Enviar a Meta Conversions API
        var capiResult = sendCapiPurchase(orderObj, paidDateStr);

        sheet.getRange(rowIndex, 15).setValue(capiResult.status);   // Col O = Estado CAPI
        sheet.getRange(rowIndex, 16).setValue(capiResult.event_id); // Col P = ID Evento CAPI

        return respondJSON({
          status: "success",
          message: "Pedido marcado como pagado.",
          id: orderIdToPay,
          payment_status: "Pagado",
          paid_at: paidDateStr,
          capi_status: capiResult.status,
          capi_event_id: capiResult.event_id,
          capi_message: capiResult.message
        });
      }
    }
    return respondJSON({ status: "error", message: "Pedido #" + orderIdToPay + " no encontrado" });
  }

  // 4. REINTENTAR CAPI PARA PEDIDO PAGADO CON FALLO PREVIO
  if (action === "retry_capi") {
    var retryId = data.id;
    var maxR = sheet.getLastRow();
    if (maxR <= 1) return respondJSON({ status: "error", message: "Hoja vacía" });

    var allRows = sheet.getRange(2, 1, maxR - 1, HEADERS.length).getValues();
    for (var m = 0; m < allRows.length; m++) {
      var rowItem = allRows[m];
      if (String(rowItem[0]).trim() === String(retryId).trim()) {
        var rowTarget = m + 2;
        var retryOrderObj = {
          id: String(rowItem[0]),
          date: String(rowItem[1]),
          name: String(rowItem[2]),
          phone: String(rowItem[3]).replace(/^'/, ''),
          dni: String(rowItem[4]).replace(/^'/, ''),
          city: String(rowItem[5]),
          address: String(rowItem[6]),
          product: String(rowItem[7]),
          units: Number(rowItem[8]) || 1,
          price: Number(rowItem[9]) || 99.90,
          payment: String(rowItem[10]),
          attribution: rowItem[16] || {}
        };

        var retryPaidDate = rowItem[13] || Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");
        var retryResult = sendCapiPurchase(retryOrderObj, retryPaidDate);

        sheet.getRange(rowTarget, 15).setValue(retryResult.status);
        sheet.getRange(rowTarget, 16).setValue(retryResult.event_id);

        return respondJSON({
          status: retryResult.status === "sent" ? "success" : "error",
          message: retryResult.message,
          id: retryId,
          capi_status: retryResult.status,
          capi_event_id: retryResult.event_id
        });
      }
    }
    return respondJSON({ status: "error", message: "Pedido no encontrado para reintento" });
  }

  // 5. REGISTRAR NUEVO PEDIDO DESDE LA LANDING
  if (action === "create_order") {
    var units = Number(data.units) || 1;
    if (!PRICE_CATALOG[units]) units = 1;
    var catalogItem = PRICE_CATALOG[units];
    var validatedPrice = catalogItem.price;
    var validatedProduct = catalogItem.name;

    var attrString = "";
    if (data.attribution) {
      attrString = typeof data.attribution === "string" ? data.attribution : JSON.stringify(data.attribution);
    }

    var orderId = data.id || ("VL-" + Math.floor(1000 + Math.random() * 9000));
    var orderDate = data.date || Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm");

    var row = [
      orderId,                                    // Col A
      orderDate,                                  // Col B
      data.name || "Sin nombre",                  // Col C
      "'" + (data.phone || ""),                   // Col D
      "'" + (data.dni || ""),                     // Col E
      data.city || "",                            // Col F
      data.address || "",                         // Col G
      data.product || validatedProduct,           // Col H
      units,                                      // Col I
      validatedPrice,                             // Col J
      data.payment || "Yape Oficial",             // Col K
      "Pendiente",                                // Col L - Estado Logístico
      "Pendiente",                                // Col M - Estado Pago (NUNCA pagado al registrar)
      "",                                         // Col N - Fecha Pago
      "none",                                     // Col O - Estado CAPI
      "",                                         // Col P - ID Evento CAPI
      attrString                                  // Col Q - Atribución
    ];

    sheet.appendRow(row);

    var newRowIdx = sheet.getLastRow();
    sheet.getRange(newRowIdx, 10).setNumberFormat('"S/" #,##0.00');

    return respondJSON({
      status: "success",
      message: "Pedido registrado con éxito",
      id: orderId
    });
  }

  // 6. ESTADO DE CONFIGURACIÓN CAPI
  if (action === "get_config_status") {
    var p = PropertiesService.getScriptProperties();
    var hasToken = !!(p.getProperty("META_CAPI_ACCESS_TOKEN") && p.getProperty("META_CAPI_ACCESS_TOKEN").trim().length > 10);
    var pId = p.getProperty("META_PIXEL_ID") || DEFAULT_PIXEL_ID;
    var testCode = p.getProperty("META_TEST_EVENT_CODE") || "";

    return respondJSON({
      status: "success",
      has_access_token: hasToken,
      pixel_id: pId,
      test_event_code: testCode
    });
  }

  // 7. GUARDAR CONFIGURACIÓN CAPI
  if (action === "set_config") {
    var scriptProps = PropertiesService.getScriptProperties();
    if (data.pixel_id) {
      scriptProps.setProperty("META_PIXEL_ID", String(data.pixel_id).trim());
    }
    if (data.test_event_code !== undefined) {
      scriptProps.setProperty("META_TEST_EVENT_CODE", String(data.test_event_code).trim());
    }
    if (data.access_token && String(data.access_token).trim() !== "") {
      scriptProps.setProperty("META_CAPI_ACCESS_TOKEN", String(data.access_token).trim());
    }

    return respondJSON({
      status: "success",
      message: "Configuración de Meta CAPI guardada con éxito en Google Apps Script"
    });
  }

  return respondJSON({ status: "error", message: "Acción no reconocida: " + action });
}

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
