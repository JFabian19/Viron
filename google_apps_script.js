/**
 * ==============================================================================
 * VELORA BEAUTY - GOOGLE APPS SCRIPT BACKEND
 * Hoja de Cálculo ID: 1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8
 * ==============================================================================
 * 
 * INSTRUCCIONES:
 * 1. Abre tu hoja de Google Sheets:
 *    https://docs.google.com/spreadsheets/d/1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8/edit
 * 2. En el menú superior: Extensiones > Apps Script.
 * 3. Borra el código actual, pega este archivo completo y guarda (Ctrl + S).
 * 4. Implementar > Nueva implementación > Tipo: "Aplicación web".
 * 5. Ejecutar como: "Yo" | Quién tiene acceso: "Cualquier persona".
 * 6. Copia la URL generada (/exec) y guárdala en tu admin.html.
 */

var SPREADSHEET_ID = "1lN4U1vUtAozuxe4v3sqqmCfXn7ZvbnoEBq35Tr_eqg8";
var SHEET_ORDERS = "Pedidos";
var SHEET_METRICS = "Métricas";

// Encabezados oficiales de la Hoja Pedidos (12 columnas limpias)
var HEADERS = [
  "ID Pedido",        // Col A
  "Fecha y Hora",     // Col B
  "Cliente",          // Col C
  "WhatsApp",         // Col D
  "DNI",              // Col E
  "Departamento",     // Col F
  "Agencia Shalom",   // Col G
  "Producto",         // Col H
  "Unidades",         // Col I
  "Total (S/)",       // Col J
  "Método Pago",      // Col K
  "Estado"            // Col L
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
 * Inicializar automáticamente hojas y encabezados
 */
function setupSheet() {
  var ss = getTargetSpreadsheet();
  
  // 1. Configurar Hoja "Pedidos"
  var ordersSheet = ss.getSheetByName(SHEET_ORDERS);
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet(SHEET_ORDERS, 0);
  }
  
  if (ordersSheet.getLastRow() === 0) {
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
  }

  // 2. Configurar Hoja "Métricas"
  var metricsSheet = ss.getSheetByName(SHEET_METRICS);
  if (!metricsSheet) {
    metricsSheet = ss.insertSheet(SHEET_METRICS, 1);
    
    metricsSheet.getRange("A1:B1").merge().setValue("📊 PANEL DE CONTROL Y MÉTRICAS VELORA")
      .setFontWeight("bold").setFontSize(14).setBackground("#0B0F19").setFontColor("#22C55E").setHorizontalAlignment("center");
    
    var metricsData = [
      ["Métrica", "Fórmula Automática"],
      ["Ventas Totales (S/)", '=IFERROR(SUM(Pedidos!J2:J), 0)'],
      ["Total de Pedidos Registrados", '=MAX(0, COUNTA(Pedidos!A2:A)-1)'],
      ["Pedidos Pendientes", '=COUNTIF(Pedidos!L2:L, "Pendiente")'],
      ["Pedidos Confirmados", '=COUNTIF(Pedidos!L2:L, "Confirmado")'],
      ["Pedidos En Camino", '=COUNTIF(Pedidos!L2:L, "En Camino")'],
      ["Pedidos Entregados", '=COUNTIF(Pedidos!L2:L, "Entregado")'],
      ["Pedidos Cancelados", '=COUNTIF(Pedidos!L2:L, "Cancelado")'],
      ["Ticket Promedio (S/)", '=IFERROR(AVERAGE(Pedidos!J2:J), 0)']
    ];
    
    metricsSheet.getRange(2, 1, metricsData.length, 2).setValues(metricsData);
    metricsSheet.getRange("A2:B2").setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
    metricsSheet.getRange("B3:B10").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right");
    metricsSheet.setColumnWidth(1, 260);
    metricsSheet.setColumnWidth(2, 200);
  }
  
  return { status: "success", message: "Hojas inicializadas con éxito" };
}

/**
 * Registrar pedidos o actualizar estados (POST)
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

    // Actualizar estado desde admin
    if (action === "update_status") {
      var targetId = data.id;
      var newStatus = data.status;
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var r = 0; r < idValues.length; r++) {
          if (String(idValues[r][0]).trim() === String(targetId).trim()) {
            sheet.getRange(r + 2, 12).setValue(newStatus); // Columna L = Estado
            return respondJSON({ status: "success", message: "Estado actualizado", id: targetId, newStatus: newStatus });
          }
        }
      }
      return respondJSON({ status: "error", message: "Pedido no encontrado" });
    }

    // Registrar nuevo pedido
    var row = [
      data.id || ("VL-" + Math.floor(1000 + Math.random() * 9000)),
      data.date || Utilities.formatDate(new Date(), "America/Lima", "yyyy-MM-dd HH:mm"),
      data.name || "Sin nombre",
      "'" + (data.phone || ""),
      "'" + (data.dni || ""),
      data.city || "",
      data.address || "",
      data.product || "1x VELORA 5 en 1",
      Number(data.units) || 1,
      Number(data.price) || 99.90,
      data.payment || "Yape Oficial",
      data.status || "Pendiente"
    ];

    sheet.appendRow(row);

    // Formato moneda en columna J
    var newRowIdx = sheet.getLastRow();
    sheet.getRange(newRowIdx, 10).setNumberFormat('"S/" #,##0.00');

    return respondJSON({
      status: "success",
      message: "Pedido guardado en Google Sheets",
      id: row[0]
    });

  } catch (error) {
    return respondJSON({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * Leer pedidos para el admin (GET)
 */
function doGet(e) {
  try {
    setupSheet();
    var ss = getTargetSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ORDERS);
    
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return respondJSON({ status: "success", orders: [] });
    }
    
    var dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    var rawValues = dataRange.getValues();
    
    var orders = [];
    for (var i = rawValues.length - 1; i >= 0; i--) {
      var row = rawValues[i];
      if (!row[0]) continue;
      
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
        status: String(row[11]) || "Pendiente"
      });
    }

    return respondJSON({
      status: "success",
      count: orders.length,
      orders: orders
    });

  } catch (error) {
    return respondJSON({
      status: "error",
      message: error.toString(),
      orders: []
    });
  }
}

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
