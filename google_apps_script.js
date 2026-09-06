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
 *        * ADMIN_SECRET           : (Obligatorio) Contraseña privada de al menos 16 caracteres para autorizar cambios sensibles desde /admin.
 * 5. Haz clic en: Implementar > Nueva implementación.
 *    - Tipo: "Aplicación web".
 *    - Ejecutar como: "Yo" (tu cuenta de Google).
 *    - Quién tiene acceso: "Cualquier persona" (Anyone).
 * 6. Copia la URL generada (/exec) y pégala en la configuración de admin.html.
 */

var SPREADSHEET_ID = "1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8";
var DEFAULT_PIXEL_ID = "1635217208321987";
var GRAPH_API_VERSION = "v26.0";
var SHEET_ORDERS = "Pedidos";
var SHEET_METRICS = "Métricas";

// Las credenciales se configuran exclusivamente en Propiedades del Script.

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
      ["Total de Pedidos Registrados", '=COUNTA(Pedidos!A2:A)'],
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

function hashExactSHA256(text) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ""), Utilities.Charset.UTF_8);
  return raw.map(function(byte) {
    var value = byte < 0 ? byte + 256 : byte;
    return value.toString(16).padStart(2, "0");
  }).join("");
}

/**
 * Enviar evento Purchase a Meta Conversions API
 */
function sendCapiPurchase(order, paidTimestamp) {
  var props = PropertiesService.getScriptProperties();
  var accessToken = props.getProperty("META_CAPI_ACCESS_TOKEN");
  var pixelId = props.getProperty("META_PIXEL_ID") || DEFAULT_PIXEL_ID;
  var testEventCode = props.getProperty("META_TEST_EVENT_CODE");

  if (!accessToken || accessToken.trim() === "") {
    return {
      status: "failed",
      event_id: "purchase_" + String(order.id).trim(),
      message: "META_CAPI_ACCESS_TOKEN no configurada en Propiedades del Script de Google Apps Script."
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
      var lastName = nameParts.slice(Math.max(1, nameParts.length - 2)).join("");
      userData.ln = [hashSHA256(lastName)];
    }
  }

  // Ciudad/Departamento (sin tildes, minúsculas)
  var cleanCity = stripAccents(order.city || "").replace(/[^a-z0-9]/g, "");
  if (cleanCity) {
    userData.st = [hashSHA256(cleanCity)];
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
  var eventEpochSeconds = Math.floor(new Date(paidTimestamp).getTime() / 1000);
  if (!isFinite(eventEpochSeconds) || eventEpochSeconds > Math.floor(Date.now() / 1000) ||
      eventEpochSeconds < Math.floor(Date.now() / 1000) - 7 * 86400) {
    return { status: "failed", event_id: eventId, message: "Fecha de pago inválida o fuera del plazo de 7 días de Meta." };
  }

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

  var url = "https://graph.facebook.com/" + GRAPH_API_VERSION + "/" + pixelId + "/events";

  try {
    var response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + accessToken.trim() },
      payload: JSON.stringify(postData),
      muteHttpExceptions: true
    });

    var resCode = response.getResponseCode();
    var resText = response.getContentText();
    var resJson = {};
    try { resJson = JSON.parse(resText); } catch(errParse) {}

    if (resCode >= 200 && resCode < 300 && Number(resJson.events_received) === 1) {
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


var API_VERSION = "2026-09-05.2";

function respondJSON(obj) {
  obj.api_version = API_VERSION;
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return respondJSON({status: "success", service: "velora-orders", pixel_id: DEFAULT_PIXEL_ID});
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "create_order";
    if (action !== "create_order") {
      var expected = PropertiesService.getScriptProperties().getProperty("ADMIN_SECRET");
      if (!expected || expected.length < 16 || hashExactSHA256(data.admin_secret || "") !== hashExactSHA256(expected)) {
        return respondJSON({status: "error", code: "unauthorized", message: "Contraseña administrativa incorrecta o sin configurar."});
      }
    }
    if (!lock.tryLock(25000)) throw new Error("Servidor ocupado. Reintenta con el mismo pedido.");
    setupSheet();
    var sheet = getTargetSpreadsheet().getSheetByName(SHEET_ORDERS);
    return respondJSON(handleRequest(action, data, sheet));
  } catch (err) {
    return respondJSON({status: "error", message: String(err.message || err)});
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function readOrders(sheet) {
  if (sheet.getLastRow() < 2) return [];
  var notes = sheet.getRange(2, 15, sheet.getLastRow() - 1, 1).getNotes();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues().map(function(r, i) {
    var attr = {};
    try { attr = JSON.parse(r[16] || "{}"); } catch (_) {}
    return {
      id: String(r[0]), date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      name: String(r[2]), phone: String(r[3]).replace(/^'/, ""), dni: String(r[4]).replace(/^'/, ""),
      city: String(r[5]), address: String(r[6]), product: String(r[7]), units: Number(r[8]),
      price: Number(r[9]), payment: String(r[10]), status: String(r[11]),
      payment_status: String(r[12] || "Pendiente"),
      paid_at: r[13] instanceof Date ? r[13].toISOString() : String(r[13] || ""),
      capi_status: String(r[14] || "none"), capi_event_id: String(r[15] || ""),
      capi_message: String(notes[i][0] || ""),
      attribution: attr, _row: i + 2
    };
  }).filter(function(o) { return o.id; });
}

function safeCell(value) {
  var text = String(value || "").trim().slice(0, 1000);
  return /^[=+\-@\t\r]/.test(text) ? "'" + text : text;
}

function normalizePaidAt(value) {
  var text = String(value || "");
  // Fechas heredadas se guardaban en Lima sin zona horaria.
  if (/^\d{4}-\d\d-\d\d \d\d:\d\d(?::\d\d)?$/.test(text)) text = text.replace(" ", "T") + "-05:00";
  return text;
}

function handleRequest(action, data, sheet) {
  var props = PropertiesService.getScriptProperties();
  if (action === "login") return {status: "success"};
  if (action === "change_password") {
    if (String(data.new_password || "").length < 16) throw new Error("Usa al menos 16 caracteres.");
    props.setProperty("ADMIN_SECRET", String(data.new_password));
    return {status: "success"};
  }
  if (action === "get_config_status") {
    return {status: "success", has_access_token: !!props.getProperty("META_CAPI_ACCESS_TOKEN"),
      pixel_id: props.getProperty("META_PIXEL_ID") || DEFAULT_PIXEL_ID,
      test_event_code: props.getProperty("META_TEST_EVENT_CODE") || "",
      message: "Token configurado no implica aceptación de eventos por Meta."};
  }
  if (action === "set_config") {
    if (data.pixel_id && String(data.pixel_id) !== DEFAULT_PIXEL_ID) throw new Error("El Pixel debe coincidir con la landing.");
    if (data.test_event_code !== undefined) props.setProperty("META_TEST_EVENT_CODE", String(data.test_event_code).trim());
    return {status: "success"};
  }
  var orders = readOrders(sheet);
  if (action === "get_orders") {
    orders.forEach(function(o) { delete o._row; });
    return {status: "success", orders: orders.reverse(), count: orders.length};
  }
  var matches = orders.filter(function(o) { return o.id === String(data.id); });
  if (matches.length > 1) throw new Error("Hay pedidos heredados con el mismo ID. Corrige sus IDs en Sheets antes de procesar el pago.");
  var order = matches[0];
  if (action === "create_order") {
    if (!/^VL-[a-zA-Z0-9-]{4,80}$/.test(String(data.id || ""))) throw new Error("ID de pedido inválido.");
    var phone = String(data.phone || "").replace(/\D/g, "").replace(/^51(?=9\d{8}$)/, "");
    if (!/^9\d{8}$/.test(phone) || !/^\d{8}$/.test(String(data.dni || "")) ||
        String(data.name || "").trim().split(/\s+/).length < 2 || !data.city || !data.address) {
      throw new Error("Datos del pedido incompletos o inválidos.");
    }
    var item = PRICE_CATALOG[Number(data.units)];
    if (!item) throw new Error("Paquete inválido.");
    if (order) {
      if (order.phone !== phone || order.dni !== String(data.dni) || order.units !== item.units) {
        throw new Error("El ID corresponde a otro pedido.");
      }
      return {status: "success", id: order.id, saved: true, duplicate: true};
    }
    var attr = data.attribution && typeof data.attribution === "object" ? data.attribution : {};
    var allowed = {};
    ["fbp","fbc","client_ip","client_user_agent","landing_url","utm_source","utm_medium","utm_campaign","utm_term","utm_content"].forEach(function(k) {
      if (attr[k]) allowed[k] = String(attr[k]).slice(0, 2000);
    });
    sheet.appendRow([String(data.id), new Date().toISOString(), safeCell(data.name), "'" + phone,
      "'" + data.dni, safeCell(data.city), safeCell(data.address), item.name, item.units,
      item.price, safeCell(data.payment || "Yape Oficial"), "Pendiente", "Pendiente", "", "none", "",
      JSON.stringify(allowed)]);
    SpreadsheetApp.flush();
    return {status: "success", saved: true, id: String(data.id)};
  }
  if (!order) throw new Error("Pedido no encontrado.");
  if (action === "update_status") {
    if (["Pendiente","Confirmado","En Camino","Entregado","Cancelado"].indexOf(data.status) < 0) throw new Error("Estado inválido.");
    sheet.getRange(order._row, 12).setValue(data.status);
    return {status: "success", id: order.id};
  }
  if (action === "mark_paid" || action === "retry_capi") {
    if ((props.getProperty("META_PIXEL_ID") || DEFAULT_PIXEL_ID) !== DEFAULT_PIXEL_ID) {
      throw new Error("META_PIXEL_ID no coincide con el Pixel de la landing.");
    }
    if (action === "retry_capi" && order.payment_status !== "Pagado") throw new Error("Solo se reintentan pedidos pagados.");
    if (order.capi_status === "sent") {
      return {status: "success", id: order.id, payment_status: order.payment_status,
        paid_at: order.paid_at, capi_status: "sent", capi_event_id: order.capi_event_id, already_sent: true};
    }
    if (order.payment_status !== "Pagado") {
      order.paid_at = new Date().toISOString();
      order.payment_status = "Pagado";
      sheet.getRange(order._row, 13, 1, 2).setValues([["Pagado", order.paid_at]]);
      SpreadsheetApp.flush();
    }
    var result = sendCapiPurchase(order, normalizePaidAt(order.paid_at));
    // Un evento de prueba nunca se cuenta como una conversión real enviada.
    var state = result.status === "sent" && props.getProperty("META_TEST_EVENT_CODE") ? "test_sent" : result.status;
    sheet.getRange(order._row, 15, 1, 2).setValues([[state, result.event_id]]);
    sheet.getRange(order._row, 15).setNote(result.message || "");
    SpreadsheetApp.flush();
    return {status: "success", id: order.id, payment_status: "Pagado", paid_at: order.paid_at,
      capi_status: state, capi_event_id: result.event_id, capi_message: result.message};
  }
  throw new Error("Acción no reconocida.");
}
