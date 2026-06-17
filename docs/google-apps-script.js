var DEEPSEEK_API_KEY = PropertiesService.getScriptProperties().getProperty("DEEPSEEK_API_KEY");
var DEFAULT_FOLDER_ID = "1u51rRx9XiKdzVzBjSg7wYgTOqmuoEjLC";
var MAX_TOOL_TURNS = 10;

// ═══════════════════════════════════════════════
//  TOOLS DEFINITION (Function Calling)
// ═══════════════════════════════════════════════

var TOOLS = [
  {
    type: "function",
    function: {
      name: "list_sheets",
      description: "Lista las hojas del simulador con nombre, filas y columnas. Usala al inicio para saber qué hojas existen y cuál modificar.",
      parameters: { type: "object", properties: {} }
    }
  },
  {
    type: "function",
    function: {
      name: "read_cells",
      description: "Lee valores reales de un rango de celdas. Usala para consultar datos del Excel. Máximo 200 filas por llamada; si necesitás más, hacé varias llamadas. El rango usa notación A1 (ej: 'B8', 'A14:J50').",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          range: { type: "string", description: "Rango en notación A1. Ej: 'B8', 'A14:J50', 'V39:V69', 'Y1:AB100'" }
        },
        required: ["sheet", "range"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_cell_info",
      description: "Obtiene información completa de una celda específica: valor mostrado, fórmula (si tiene) y nota/comentario. Ideal para celdas de la fila 8 (capital, tasas, fechas) donde las fórmulas son relevantes.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          cell: { type: "string", description: "Referencia de celda en notación A1. Ej: 'B8', 'F8', 'Q8'" }
        },
        required: ["sheet", "cell"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_cells",
      description: "Escribe valores en un rango de celdas. Usala para modificar: abonos, cambiar tasas, fechas, capital. Para limpiar un rango, usá valores vacíos (''). Fechas en string 'dd/mm/yyyy'. Números sin formato ni separadores de miles. Tasas en decimal (15% = 0.15).",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          range: { type: "string", description: "Rango en notación A1. Ej: 'F8', 'Y42:AB42', 'B4'" },
          values: {
            type: "array",
            items: { type: "array", items: {} },
            description: "Array de filas con valores. Ej para 1 celda: [[200]]. Ej para fila de abono: [[200,0,0,200]]"
          }
        },
        required: ["sheet", "range", "values"]
      }
    }
  }
];

// ═══════════════════════════════════════════════
//  ENDPOINTS HTTP
// ═══════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === "chat") return handleChat(data);

    // Upload Excel
    var archivoBase64 = data.fileBase64;
    var nombreArchivo = data.fileName;
    var folderId = data.folderId || DEFAULT_FOLDER_ID;

    if (!archivoBase64 || !nombreArchivo) {
      return jsonResponse({ success: false, error: "Falta fileBase64 o fileName." });
    }

    // Validar carpeta
    try {
      var folder = DriveApp.getFolderById(folderId);
    } catch (e) {
      return jsonResponse({
        success: false,
        error: "Carpeta Drive no encontrada o sin acceso. Verificá el ID: " + folderId
      });
    }
    var decoded = Utilities.base64Decode(archivoBase64);
    var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);
    var excelFile = folder.createFile(blob);

    var metadata = {
      name: nombreArchivo.replace(".xlsx", "").replace(".xls", ""),
      parents: [folderId],
      mimeType: "application/vnd.google-apps.spreadsheet"
    };

    var googleSheet = Drive.Files.create(metadata, excelFile.getBlob());
    excelFile.setTrashed(true);

    var sheetFile = DriveApp.getFileById(googleSheet.id);
    sheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

    var urlEditor = "https://docs.google.com/spreadsheets/d/" + googleSheet.id + "/edit?usp=sharing";
    return jsonResponse({ success: true, id: googleSheet.id, url: urlEditor });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

function doGet(e) {
  try {
    if (e.parameter.action === "delete" && e.parameter.id) {
      DriveApp.getFileById(e.parameter.id).setTrashed(true);
      return jsonResponse({ success: true });
    }
    return jsonResponse({ success: false, error: "Acción no válida" });
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ═══════════════════════════════════════════════
//  CHAT CON FUNCTION CALLING
// ═══════════════════════════════════════════════

function handleChat(data) {
  try {
    if (!DEEPSEEK_API_KEY) {
      return jsonResponse({ success: false, error: "API Key no configurada." });
    }

    var sheetId = data.id;
    var question = data.question;
    var history = data.history || [];

    if (!sheetId || !question) {
      return jsonResponse({ success: false, error: "Falta id del sheet o question." });
    }

    var ss = SpreadsheetApp.openById(sheetId);

    // Construir mensajes iniciales
    var messages = [{ role: "system", content: buildSystemPrompt() }];

    for (var i = 0; i < history.length; i++) {
      if (history[i] && history[i].role && history[i].content && history[i].role !== "system") {
        messages.push({ role: history[i].role, content: history[i].content });
      }
    }
    messages.push({ role: "user", content: question });

    // Loop de tool calling
    var turn = 0;
    var finalReply = "";
    var lastContent = "";

    while (turn < MAX_TOOL_TURNS) {
      turn++;

      var payload = {
        model: "deepseek-chat",
        messages: messages,
        tools: TOOLS,
        temperature: 0.1
      };

      var options = {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + DEEPSEEK_API_KEY },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch("https://api.deepseek.com/v1/chat/completions", options);
      var result = JSON.parse(response.getContentText());

      if (result.error) {
        return jsonResponse({ success: false, error: "DeepSeek: " + (result.error.message || "Error desconocido") });
      }

      var msg = result.choices[0].message;
      lastContent = msg.content || "";
      console.log("Turn " + turn + ": finish=" + result.choices[0].finish_reason + ", content=" + lastContent.length + " chars, tool_calls=" + (msg.tool_calls ? msg.tool_calls.length : 0));

      // Si el modelo quiere llamar herramientas
      if (msg.tool_calls && msg.tool_calls.length > 0) {
        // Agregar respuesta del asistente con tool_calls al historial
        messages.push({
          role: "assistant",
          content: msg.content || null,
          tool_calls: msg.tool_calls
        });

        // Ejecutar cada tool call
        for (var t = 0; t < msg.tool_calls.length; t++) {
          var tc = msg.tool_calls[t];
          var fnName = tc.function.name;
          // arguments puede venir como string JSON o como objeto ya parseado
          var fnArgs = tc.function.arguments;
          if (typeof fnArgs === "string") fnArgs = safeJsonParse(fnArgs);
          if (!fnArgs) fnArgs = {};

          console.log("  Tool [" + turn + "]: " + fnName + "(" + JSON.stringify(fnArgs).substring(0, 100) + ")");

          var toolResult = executeTool(ss, fnName, fnArgs);
          console.log("  Result: " + toolResult.substring(0, 150));

          messages.push({
            role: "tool",
            tool_call_id: tc.id,
            content: toolResult
          });
        }
        continue; // Siguiente turno
      }

      // Sin tool_calls → respuesta final
      finalReply = lastContent;
      break;
    }

    if (!finalReply && lastContent) {
      // Si hay contenido parcial usarlo
      finalReply = lastContent;
    }

    if (!finalReply) {
      return jsonResponse({
        success: false,
        error: "El asistente no generó respuesta después de " + turn + " turnos. Intentá una pregunta más concreta."
      });
    }

    return jsonResponse({ success: true, reply: finalReply });

  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ═══════════════════════════════════════════════
//  EJECUCIÓN DE TOOLS
// ═══════════════════════════════════════════════

function executeTool(ss, toolName, args) {
  try {
    if (toolName === "list_sheets") {
      return toolListSheets(ss);
    }
    if (toolName === "get_cell_info") {
      return toolGetCellInfo(ss, args.sheet, args.cell);
    }
    if (toolName === "read_cells") {
      return toolReadCells(ss, args.sheet, args.range);
    }
    if (toolName === "write_cells") {
      return toolWriteCells(ss, args.sheet, args.range, args.values);
    }
    return "ERROR: herramienta desconocida: " + toolName;
  } catch (e) {
    return "ERROR: " + e.toString();
  }
}

function toolListSheets(ss) {
  var sheets = ss.getSheets();
  var result = [];
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var name = s.getName();
    var nameNorm = normalizeText(name);
    // Omitir hojas de referencia
    if (nameNorm === "base" || nameNorm === "obligaciones" || nameNorm.indexOf("tasa usura") !== -1) continue;
    result.push({
      name: name,
      rows: s.getLastRow(),
      cols: s.getLastColumn()
    });
  }
  if (result.length === 0) return "No hay hojas de simulación (se omitieron Base, Obligaciones y Tasa usura).";
  return JSON.stringify(result);
}

function toolGetCellInfo(ss, sheetName, cellRef) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";

  var parsed = parseRange(cellRef);
  if (!parsed || parsed.numRows !== 1 || parsed.numCols !== 1) {
    return "ERROR: '" + cellRef + "' no es una celda válida. Usá formato A1 (ej: 'B8', 'F8').";
  }

  var range = sheet.getRange(parsed.startRow, parsed.startCol);
  var value = range.getValue();
  var formula = range.getFormula();
  var note = "";
  try { note = range.getNote() || ""; } catch (e) { /* algunas celdas no tienen notas */ }

  var lines = [];
  lines.push("Celda: " + cellRef);

  if (formula) {
    lines.push("Fórmula: " + formula);
    lines.push("Valor calculado: " + formatValue(value));
  } else {
    lines.push("Valor: " + formatValue(value));
  }

  if (note) lines.push("Nota: " + note);

  return lines.join("\n");
}

function toolReadCells(ss, sheetName, rangeStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";

  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango inválido '" + rangeStr + "'. Usá notación A1 (ej: 'B8', 'A14:J50').";

  // Limitar tamaño máximo
  var numRows = parsed.numRows;
  var numCols = parsed.numCols;
  if (numRows > 200) {
    numRows = 200;
    parsed.endRow = parsed.startRow + 199;
  }

  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);
  var values = range.getValues();
  var formulas = range.getFormulas();

  // Caso celda única
  if (numRows === 1 && numCols === 1) {
    var v = values[0][0];
    var f = formulas[0][0];
    if (f) return formatValue(v) + " (fórmula: " + f + ")";
    return formatValue(v);
  }

  // Caso rango: devolver filas con formato compacto
  var rows = [];
  var nonEmptyCols = [];
  // Encontrar columnas con al menos un valor no vacío
  for (var c = 0; c < numCols; c++) {
    var hasVal = false;
    for (var r = 0; r < numRows; r++) {
      if (values[r][c] !== "" && values[r][c] !== null && values[r][c] !== undefined) { hasVal = true; break; }
      if (formulas[r][c]) { hasVal = true; break; }
    }
    nonEmptyCols.push(hasVal);
  }

  for (var r = 0; r < numRows; r++) {
    var cells = [];
    for (var c = 0; c < numCols; c++) {
      if (!nonEmptyCols[c]) continue;
      var val = values[r][c];
      var formula = formulas[r][c];
      if (formula || (val !== "" && val !== null && val !== undefined)) {
        cells.push(columnToLetter(parsed.startCol - 1 + c) + "=" + formatValue(val));
      }
    }
    if (cells.length > 0) {
      rows.push("r" + (parsed.startRow + r) + ":" + cells.join(" | "));
    }
  }

  if (rows.length === 0) return "(rango vacío)";
  return rows.join("\n") + (numRows < parsed.numRows ? "\n[Truncado a 200 filas]" : "");
}

function toolWriteCells(ss, sheetName, rangeStr, values) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";

  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango inválido '" + rangeStr + "'.";

  if (!values || !values.length || !values[0].length) {
    return "ERROR: values vacío o inválido.";
  }

  // Normalizar valores (fechas, comas decimales)
  var normalized = [];
  for (var r = 0; r < values.length; r++) {
    var row = [];
    for (var c = 0; c < values[r].length; c++) {
      row.push(normalizeInputValue(values[r][c]));
    }
    normalized.push(row);
  }

  var numRows = Math.min(normalized.length, parsed.numRows);
  var numCols = Math.min(normalized[0].length, parsed.numCols);

  // Recortar al tamaño exacto del rango pedido
  var writeValues = [];
  for (var rr = 0; rr < numRows; rr++) {
    var row = [];
    for (var cc = 0; cc < numCols; cc++) {
      row.push(normalized[rr][cc]);
    }
    writeValues.push(row);
  }

  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);

  // Si todos los valores son vacíos, usar clearContent
  var allEmpty = true;
  for (var re = 0; re < writeValues.length && allEmpty; re++) {
    for (var ce = 0; ce < writeValues[re].length && allEmpty; ce++) {
      if (writeValues[re][ce] !== "" && writeValues[re][ce] !== null && writeValues[re][ce] !== undefined) allEmpty = false;
    }
  }

  if (allEmpty) {
    range.clearContent();
    SpreadsheetApp.flush();
    return "✓ Limpiado " + sheetName + "!" + rangeStr;
  }

  range.setValues(writeValues);
  SpreadsheetApp.flush();
  return "✓ Escrito en " + sheetName + "!" + rangeStr + " → " + JSON.stringify(writeValues).substring(0, 200);
}

// ═══════════════════════════════════════════════
//  PARSEO DE RANGOS A1
// ═══════════════════════════════════════════════

function parseRange(rangeStr) {
  var match = rangeStr.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!match) return null;

  var startCol = colLetterToNumber(match[1]);
  var startRow = parseInt(match[2], 10);

  if (match[3]) {
    var endCol = colLetterToNumber(match[3]);
    var endRow = parseInt(match[4], 10);
    return {
      startCol: startCol,
      startRow: startRow,
      endCol: endCol,
      endRow: endRow,
      numRows: endRow - startRow + 1,
      numCols: endCol - startCol + 1
    };
  }

  return {
    startCol: startCol,
    startRow: startRow,
    endCol: startCol,
    endRow: startRow,
    numRows: 1,
    numCols: 1
  };
}

function colLetterToNumber(letter) {
  var col = 0;
  letter = letter.toUpperCase();
  for (var i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

// ═══════════════════════════════════════════════
//  PROMPT DEL SISTEMA
// ═══════════════════════════════════════════════

function buildSystemPrompt() {
  return ""
    + "Eres un asistente experto en análisis y modificación de un simulador financiero en Google Sheets.\n\n"

    + "USÁ LAS HERRAMIENTAS DISPONIBLES:\n"
    + "- list_sheets: primero para saber qué hojas existen.\n"
    + "- get_cell_info: para inspeccionar una celda específica (valor + fórmula + nota). Ideal para F8, B8, Q8, etc.\n"
    + "- read_cells: para leer rangos de datos. Leé SOLO lo que necesitás, no pidas más de 50 filas por vez.\n"
    + "- write_cells: para modificar datos (abonos, tasas, fechas, capital).\n\n"

    + "ESTRUCTURA TÍPICA DE CADA HOJA DE SIMULACIÓN:\n"
    + "Fila 1: encabezados de columnas.\n"
    + "Filas 2-7: datos del pagador, proveedor, fechas.\n"
    + "Fila 8: valores principales de la obligación:\n"
    + "  A=Obligación | B=Capital | C=Fecha Transferencia | D=Vcto Neto | E=Vcto Total\n"
    + "  F=Tasa Negocio EA% | L=Tasa Int Remuneratorio EA% | AB=Tasa DPP\n"
    + "  Q=Saldo Capital | Z=Saldo Int Rem | AA=Saldo Int Mora\n"
    + "  AD=Parámetro GMF | AG=Rte Fte | AH=Valor Giro\n"
    + "Fila 13: encabezados de tablas diarias.\n"
    + "Filas 14+: datos diarios.\n\n"

    + "TABLAS DIARIAS (columnas aproximadas, leé los headers reales con read_cells):\n"
    + "1) Amortización (A:J): Fecha, Días, Saldo Capital, Amortización Diaria, etc.\n"
    + "2) Interés Remuneratorio (L:Q): Tasa, Días, Causación diaria, Acumulado.\n"
    + "3) Interés Moratorio (S:W): Tasa, Días, Causación diaria, Acumulado.\n"
    + "4) Abonos (Y:AB): Abono a Capital, Abono a Int Rem, Abono a Int Mora, Total Abono.\n\n"

    + "MODIFICACIONES:\n"
    + "- Para un abono: primero leé la tabla Y:AB con read_cells para encontrar la fila correcta por fecha.\n"
    + "  Luego escribí con write_cells en esa fila: [abonoCap, abonoIntRem, abonoIntMora, total].\n"
    + "- Para cambiar tasa/capital/fecha: write_cells en la celda específica.\n"
    + "- Fechas: string 'dd/mm/yyyy'.\n"
    + "- Tasas: decimal (15% = 0.15).\n"
    + "- Montos: número sin separadores de miles.\n"
    + "- Para limpiar abonos: write_cells con valores vacíos ('').\n\n"

    + "REGLAS:\n"
    + "- NUNCA inventes valores. Si no leíste un dato con read_cells, no lo asumas.\n"
    + "- Si un dato no está en el Excel, decilo.\n"
    + "- Para sumatorias o cálculos: leé las celdas necesarias con read_cells y luego calculá.\n"
    + "- Para mostrar datos tabulares usá formato tabla markdown:\n"
    + "  | Fecha | Valor |\n"
    + "  |-------|-------|\n"
    + "  | 10/05 | 208,45 |\n"
    + "  | 11/05 | 208,45 |\n"
    + "- Responde en español, conciso.\n"
    + "- Al finalizar una modificación, indicá claramente qué cambiaste y en qué celda.\n\n"
    + "ESTILO DE RESPUESTA (OBLIGATORIO):\n"
    + "- NUNCA uses saludos ni despedidas (nada de 'Listo.', 'Perfecto.', 'Claro.', 'Entendido.').\n"
    + "- NUNCA pidas disculpas ni digas 'disculpe', 'perdón', 'tiene razón'.\n"
    + "- NUNCA repitas la pregunta del usuario ni expliques lo obvio.\n"
    + "- NUNCA digas 'Aquí tienes', 'Aquí está', 'Te muestro'.\n"
    + "- Ve DIRECTAMENTE al resultado: responde con la tabla, el número o el dato solicitado.\n"
    + "- Si la respuesta es solo un número, respondé solo el número.\n"
    + "- Si hay una tabla, mostrala sin texto introductorio ni explicaciones innecesarias.\n"
    + "- Solo agregá una explicación breve si el usuario la pide explícitamente o si el resultado requiere contexto.\n"
    + "- Máximo 1 oración de contexto antes de una tabla. Si la tabla habla por sí sola, cero oraciones.";
}

// ═══════════════════════════════════════════════
//  UTILIDADES
// ═══════════════════════════════════════════════

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function safeJsonParse(text) {
  try { return JSON.parse(text); } catch (e) { return null; }
}

function normalizeText(text) {
  return String(text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
}

function normalizeInputValue(value) {
  if (typeof value === "string") {
    var t = value.trim();
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(t)) {
      var parts = t.split("/");
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
    if (/^-?\d+,\d+$/.test(t)) {
      return parseFloat(t.replace(",", "."));
    }
  }
  return value;
}

function findSheet(ss, name) {
  if (!name) return null;
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var norm = normalizeText(name);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalizeText(sheets[i].getName()) === norm) return sheets[i];
  }
  for (var j = 0; j < sheets.length; j++) {
    if (normalizeText(sheets[j].getName()).indexOf(norm) !== -1) return sheets[j];
  }
  return null;
}

function formatValue(val) {
  if (val === "" || val === null || val === undefined) return "vacio";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  if (typeof val === "number") {
    if (val === Math.floor(val)) return val.toString();
    return val.toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 6 });
  }
  return val.toString();
}

function columnToLetter(col) {
  var letter = "";
  while (col >= 0) {
    letter = String.fromCharCode((col % 26) + 65) + letter;
    col = Math.floor(col / 26) - 1;
  }
  return letter;
}
