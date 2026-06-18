var DEEPSEEK_API_KEY = PropertiesService.getScriptProperties().getProperty("DEEPSEEK_API_KEY");
var DEFAULT_FOLDER_ID = "1u51rRx9XiKdzVzBjSg7wYgTOqmuoEjLC";
var MAX_TOOL_TURNS = 24;
var DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions";
var DEEPSEEK_MODEL = "deepseek-chat";
var PROTECTED_SHEETS = ["base", "obligaciones", "tasa usura"];
var MAX_READ_ROWS = 200;

// ═══════════════════════════════════════════════
//  TOOLS DEFINITION (Function Calling)
// ═══════════════════════════════════════════════

var TOOLS = [
  {
    type: "function",
    function: {
      name: "list_sheets",
      description: "Lista las hojas del archivo con nombre, filas, columnas y tipo. Usala al inicio. Base, Obligaciones y Tasa usura son protegidas; el contexto estructural sale SOLO de Base.",
      parameters: { type: "object", properties: {} }
    }
  },
  {
    type: "function",
    function: {
      name: "get_simulator_schema",
      description: "Construye un mapa dinámico de la hoja leyendo encabezados reales: campos principales, tabla diaria, columnas de fecha, intereses y abonos. Usa la estructura de Base como referencia.",
      parameters: {
        type: "object",
        properties: { sheet: { type: "string", description: "Nombre exacto de la hoja" } },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_summary",
      description: "Resume los campos principales de la obligación: obligación, capital, fechas, tasas, saldos, GMF, Rte Fte y Valor Giro.",
      parameters: {
        type: "object",
        properties: { sheet: { type: "string", description: "Nombre exacto de la hoja" } },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "find_date_row",
      description: "Busca una fecha dd/mm/yyyy en la tabla diaria. Si no se envía fecha, devuelve la última fecha útil.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          date: { type: "string", description: "Fecha dd/mm/yyyy. Opcional." }
        },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "apply_payment",
      description: "Aplica un abono en una sola llamada. Encuentra la fila por fecha y las columnas de abono por encabezado: capital, interés remuneratorio, interés moratorio y total.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          date: { type: "string", description: "Fecha dd/mm/yyyy. Opcional; si falta, usa última fecha útil." },
          capital: { type: "number", description: "Abono a capital. Si el usuario no discrimina, usar aquí el monto." },
          interestRem: { type: "number", description: "Abono a interés remuneratorio. Default 0." },
          interestMora: { type: "number", description: "Abono a interés moratorio. Default 0." }
        },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "clear_payment",
      description: "Limpia las columnas de abono de una fecha: capital, interés remuneratorio, interés moratorio y total.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          date: { type: "string", description: "Fecha dd/mm/yyyy. Opcional; si falta, usa última fecha útil." }
        },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_named_field",
      description: "Modifica un campo principal por nombre lógico: capital, fecha_transferencia, fecha_vcto_neto, fecha_vcto_total, tasa_negocio, tasa_remuneratorio, tasa_dpp, parametro_gmf, parametro_rte_fte.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          field: { type: "string" },
          value: { description: "Valor a escribir. Fechas dd/mm/yyyy; tasas en decimal o porcentaje; montos como número." }
        },
        required: ["sheet", "field", "value"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "read_cells",
      description: "Lee valores reales de un rango de celdas. Máximo 200 filas por llamada. Rango A1: 'B8', 'A14:J50', 'Y1:AB100'.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          range: { type: "string", description: "Rango en notación A1" }
        },
        required: ["sheet", "range"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_cell_info",
      description: "Obtiene valor, fórmula, nota y encabezado probable de una celda específica.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          cell: { type: "string", description: "Celda A1. Ej: 'B8', 'F8', 'Q8'" }
        },
        required: ["sheet", "cell"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_cells",
      description: "Escribe valores en un rango A1. Usar solo si no existe una tool específica. Para limpiar, usar valores vacíos ''.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          range: { type: "string", description: "Rango A1" },
          values: { type: "array", items: { type: "array", items: {} } }
        },
        required: ["sheet", "range", "values"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "calculate",
      description: "Ejecuta una fórmula en la hoja objetivo. USALA SIEMPRE para sumatorias, promedios, restas, conteos y cualquier cálculo. Las referencias sin calificar apuntan a la hoja indicada.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string", description: "Nombre exacto de la hoja" },
          formula: { type: "string", description: "Fórmula completa con =. Ej: '=SUM(V106:V181)'" }
        },
        required: ["sheet", "formula"]
      }
    }
  }
];

// ═══════════════════════════════════════════════
//  ENDPOINTS HTTP
// ═══════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || "{}");
    if (data.action === "chat") return handleChat(data);

    var archivoBase64 = data.fileBase64;
    var nombreArchivo = data.fileName;
    var folderId = data.folderId || DEFAULT_FOLDER_ID;

    if (!archivoBase64 || !nombreArchivo) {
      return jsonResponse({ success: false, error: "Falta fileBase64 o fileName." });
    }

    try {
      var folder = DriveApp.getFolderById(folderId);
    } catch (errFolder) {
      return jsonResponse({ success: false, error: "Carpeta Drive no encontrada o sin acceso. ID: " + folderId });
    }

    var decoded = Utilities.base64Decode(archivoBase64);
    var blob = Utilities.newBlob(decoded, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreArchivo);
    var excelFile = folder.createFile(blob);

    var metadata = {
      name: nombreArchivo.replace(/\.xlsx?$/i, ""),
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
    if (!DEEPSEEK_API_KEY) return jsonResponse({ success: false, error: "API Key no configurada." });

    var sheetId = data.id;
    var question = data.question;
    var history = data.history || [];

    if (!sheetId || !question) return jsonResponse({ success: false, error: "Falta id del sheet o question." });

    var ss = SpreadsheetApp.openById(sheetId);
    var messages = [{ role: "system", content: buildSystemPrompt() }];

    for (var i = 0; i < history.length; i++) {
      var h = history[i];
      if (h && h.role && h.content && h.role !== "system") {
        messages.push({ role: h.role, content: String(h.content).slice(0, 12000) });
      }
    }
    messages.push({ role: "user", content: question });

    var turn = 0;
    var finalReply = "";
    var lastContent = "";

    while (turn < MAX_TOOL_TURNS) {
      turn++;

      var payload = { model: DEEPSEEK_MODEL, messages: messages, tools: TOOLS, temperature: 0.05 };
      var options = {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + DEEPSEEK_API_KEY },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(DEEPSEEK_URL, options);
      var raw = response.getContentText();
      var result = safeJsonParse(raw);

      if (!result) return jsonResponse({ success: false, error: "DeepSeek no devolvió JSON válido: " + raw.substring(0, 500) });
      if (result.error) return jsonResponse({ success: false, error: "DeepSeek: " + (result.error.message || JSON.stringify(result.error)) });
      if (!result.choices || !result.choices[0] || !result.choices[0].message) {
        return jsonResponse({ success: false, error: "Respuesta inesperada de DeepSeek: " + raw.substring(0, 800) });
      }

      var msg = result.choices[0].message;
      lastContent = msg.content || "";
      console.log("Turn " + turn + ": finish=" + result.choices[0].finish_reason + ", content=" + lastContent.length + ", tool_calls=" + (msg.tool_calls ? msg.tool_calls.length : 0));

      if (msg.tool_calls && msg.tool_calls.length > 0) {
        messages.push({ role: "assistant", content: msg.content || null, tool_calls: msg.tool_calls });
        for (var t = 0; t < msg.tool_calls.length; t++) {
          var tc = msg.tool_calls[t];
          var fnName = tc.function.name;
          var fnArgs = tc.function.arguments;
          if (typeof fnArgs === "string") fnArgs = safeJsonParse(fnArgs);
          if (!fnArgs) fnArgs = {};

          console.log("  Tool [" + turn + "]: " + fnName + "(" + JSON.stringify(fnArgs).substring(0, 150) + ")");
          var toolResult = executeTool(ss, fnName, fnArgs);
          console.log("  Result: " + String(toolResult).substring(0, 300));

          messages.push({ role: "tool", tool_call_id: tc.id, content: String(toolResult).slice(0, 18000) });
        }
        continue;
      }

      finalReply = lastContent;
      break;
    }

    if (!finalReply && lastContent) finalReply = lastContent;
    if (!finalReply) return jsonResponse({ success: false, error: "El asistente no generó respuesta después de " + turn + " turnos. Intentá una pregunta más concreta." });

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
    if (toolName === "list_sheets") return toolListSheets(ss);
    if (toolName === "get_simulator_schema") return toolGetSimulatorSchema(ss, args.sheet);
    if (toolName === "get_summary") return toolGetSummary(ss, args.sheet);
    if (toolName === "find_date_row") return toolFindDateRow(ss, args.sheet, args.date);
    if (toolName === "apply_payment") return toolApplyPayment(ss, args.sheet, args.date, args.capital, args.interestRem, args.interestMora);
    if (toolName === "clear_payment") return toolClearPayment(ss, args.sheet, args.date);
    if (toolName === "write_named_field") return toolWriteNamedField(ss, args.sheet, args.field, args.value);
    if (toolName === "get_cell_info") return toolGetCellInfo(ss, args.sheet, args.cell);
    if (toolName === "read_cells") return toolReadCells(ss, args.sheet, args.range);
    if (toolName === "write_cells") return toolWriteCells(ss, args.sheet, args.range, args.values);
    if (toolName === "calculate") return toolCalculate(ss, args.sheet, args.formula);
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
    result.push({
      name: name,
      type: isProtectedSheet(name) ? "protegida" : "simulacion",
      rows: s.getLastRow(),
      cols: s.getLastColumn()
    });
  }
  return JSON.stringify(result);
}

function toolGetSimulatorSchema(ss, sheetName) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  return JSON.stringify(buildSchema(sheet));
}

function toolGetSummary(ss, sheetName) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  var out = { hoja: sheet.getName(), campos: {}, ultima_fecha_diaria: null };
  var keys = Object.keys(schema.fields);
  for (var i = 0; i < keys.length; i++) {
    var f = schema.fields[keys[i]];
    var rg = sheet.getRange(f.cell);
    out.campos[keys[i]] = { label: f.label, cell: f.cell, value: formatValue(rg.getValue()), formula: rg.getFormula() || "" };
  }
  if (schema.daily.dateCol) {
    var last = findLastDateRow(sheet, schema.daily.dateCol, schema.daily.dataStartRow, sheet.getLastRow());
    out.ultima_fecha_diaria = last ? { row: last, fecha: formatValue(sheet.getRange(last, schema.daily.dateCol).getValue()) } : null;
  }
  return JSON.stringify(out);
}

function toolFindDateRow(ss, sheetName, dateStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  if (!schema.daily.dateCol) return "ERROR: no encontré columna Fecha en la tabla diaria.";
  var row = dateStr ? findRowByDate(sheet, schema.daily.dateCol, schema.daily.dataStartRow, sheet.getLastRow(), dateStr) : findLastDateRow(sheet, schema.daily.dateCol, schema.daily.dataStartRow, sheet.getLastRow());
  if (!row) return "ERROR: fecha no encontrada: " + (dateStr || "última fecha útil");
  return JSON.stringify(rowSnapshot(sheet, schema, row));
}

function toolApplyPayment(ss, sheetName, dateStr, capital, interestRem, interestMora) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura). Usá una hoja de simulación.";

  var schema = buildSchema(sheet);
  var d = schema.daily;
  if (!d.dateCol || !d.abonoCapitalCol || !d.abonoRemCol || !d.abonoMoraCol || !d.totalAbonoCol) {
    return "ERROR: no encontré columnas completas de Abonos. Ejecutá get_simulator_schema.";
  }

  var row = dateStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow(), dateStr) : findLastDateRow(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow());
  if (!row) return "ERROR: fecha de abono no encontrada: " + (dateStr || "última fecha útil");

  var cap = toNumberOrZero(capital);
  var rem = toNumberOrZero(interestRem);
  var mora = toNumberOrZero(interestMora);
  var total = cap + rem + mora;

  var startCol = Math.min(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol, d.totalAbonoCol);
  var endCol = Math.max(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol, d.totalAbonoCol);
  var width = endCol - startCol + 1;
  var vals = sheet.getRange(row, startCol, 1, width).getValues()[0];
  vals[d.abonoCapitalCol - startCol] = cap;
  vals[d.abonoRemCol - startCol] = rem;
  vals[d.abonoMoraCol - startCol] = mora;
  vals[d.totalAbonoCol - startCol] = total;

  sheet.getRange(row, startCol, 1, width).setValues([vals]);
  SpreadsheetApp.flush();

  return JSON.stringify({
    ok: true,
    accion: "abono_aplicado",
    hoja: sheet.getName(),
    fila: row,
    fecha: formatValue(sheet.getRange(row, d.dateCol).getValue()),
    celdas: {
      capital: columnToLetter(d.abonoCapitalCol) + row,
      interes_remuneratorio: columnToLetter(d.abonoRemCol) + row,
      interes_moratorio: columnToLetter(d.abonoMoraCol) + row,
      total: columnToLetter(d.totalAbonoCol) + row
    },
    valores: { capital: cap, interes_remuneratorio: rem, interes_moratorio: mora, total: total }
  });
}

function toolClearPayment(ss, sheetName, dateStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var schema = buildSchema(sheet);
  var d = schema.daily;
  if (!d.dateCol || !d.abonoCapitalCol || !d.abonoRemCol || !d.abonoMoraCol || !d.totalAbonoCol) return "ERROR: no encontré columnas de Abonos.";
  var row = dateStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow(), dateStr) : findLastDateRow(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow());
  if (!row) return "ERROR: fecha no encontrada: " + (dateStr || "última fecha útil");
  var startCol = Math.min(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol, d.totalAbonoCol);
  var endCol = Math.max(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol, d.totalAbonoCol);
  sheet.getRange(row, startCol, 1, endCol - startCol + 1).clearContent();
  SpreadsheetApp.flush();
  return "✓ Limpiado " + sheet.getName() + "!" + columnToLetter(startCol) + row + ":" + columnToLetter(endCol) + row;
}

function toolWriteNamedField(ss, sheetName, field, value) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var schema = buildSchema(sheet);
  var key = normalizeFieldAlias(field);
  var meta = schema.fields[key];
  if (!meta) return "ERROR: campo no encontrado: " + field + ". Disponibles: " + Object.keys(schema.fields).join(", ");
  var v = normalizeInputValue(value, key);
  sheet.getRange(meta.cell).setValue(v);
  SpreadsheetApp.flush();
  return JSON.stringify({ ok: true, hoja: sheet.getName(), campo: key, label: meta.label, celda: meta.cell, valor: formatValue(sheet.getRange(meta.cell).getValue()) });
}

function toolGetCellInfo(ss, sheetName, cellRef) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var parsed = parseRange(cellRef);
  if (!parsed || parsed.numRows !== 1 || parsed.numCols !== 1) return "ERROR: '" + cellRef + "' no es una celda válida. Usá formato A1.";
  var range = sheet.getRange(parsed.startRow, parsed.startCol);
  var value = range.getValue();
  var formula = range.getFormula();
  var note = "";
  try { note = range.getNote() || ""; } catch (e) {}
  var schema = buildSchema(sheet);
  var header = inferHeaderForCell(sheet, schema, parsed.startRow, parsed.startCol);
  var lines = ["Celda: " + cellRef, "Encabezado probable: " + (header || "vacio")];
  if (formula) { lines.push("Fórmula: " + formula); lines.push("Valor calculado: " + formatValue(value)); }
  else { lines.push("Valor: " + formatValue(value)); }
  if (note) lines.push("Nota: " + note);
  return lines.join("\n");
}

function toolReadCells(ss, sheetName, rangeStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango inválido '" + rangeStr + "'. Usá notación A1.";
  var numRows = parsed.numRows;
  var numCols = parsed.numCols;
  if (numRows > MAX_READ_ROWS) { numRows = MAX_READ_ROWS; parsed.endRow = parsed.startRow + MAX_READ_ROWS - 1; }
  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);
  var values = range.getValues();
  var formulas = range.getFormulas();
  if (numRows === 1 && numCols === 1) {
    var v = values[0][0], f = formulas[0][0];
    return f ? formatValue(v) + " (fórmula: " + f + ")" : formatValue(v);
  }
  var rows = [];
  var nonEmptyCols = [];
  for (var c = 0; c < numCols; c++) {
    var hasVal = false;
    for (var r = 0; r < numRows; r++) {
      if (!isBlank(values[r][c]) || formulas[r][c]) { hasVal = true; break; }
    }
    nonEmptyCols.push(hasVal);
  }
  for (var rr = 0; rr < numRows; rr++) {
    var cells = [];
    for (var cc = 0; cc < numCols; cc++) {
      if (!nonEmptyCols[cc]) continue;
      var val = values[rr][cc], formula = formulas[rr][cc];
      if (!isBlank(val) || formula) cells.push(columnToLetter(parsed.startCol + cc) + "=" + formatValue(val) + (formula ? " {f=" + formula + "}" : ""));
    }
    if (cells.length > 0) rows.push("r" + (parsed.startRow + rr) + ":" + cells.join(" | "));
  }
  if (rows.length === 0) return "(rango vacío)";
  return rows.join("\n") + (numRows < parsed.numRows ? "\n[Truncado a 200 filas]" : "");
}

function toolWriteCells(ss, sheetName, rangeStr, values) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango inválido '" + rangeStr + "'.";
  if (!values || !values.length || !values[0].length) return "ERROR: values vacío o inválido.";

  var normalized = [];
  for (var r = 0; r < values.length; r++) {
    var row = [];
    for (var c = 0; c < values[r].length; c++) row.push(normalizeInputValue(values[r][c]));
    normalized.push(row);
  }

  var numRows = Math.min(normalized.length, parsed.numRows);
  var numCols = Math.min(normalized[0].length, parsed.numCols);
  var writeValues = [];
  for (var rr = 0; rr < numRows; rr++) {
    var rowOut = [];
    for (var cc = 0; cc < numCols; cc++) rowOut.push(normalized[rr][cc]);
    writeValues.push(rowOut);
  }
  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);
  if (allEmpty(writeValues)) { range.clearContent(); SpreadsheetApp.flush(); return "✓ Limpiado " + sheetName + "!" + rangeStr; }
  range.setValues(writeValues);
  SpreadsheetApp.flush();
  return "✓ Escrito en " + sheetName + "!" + rangeStr + " → " + JSON.stringify(writeValues).substring(0, 200);
}

function toolCalculate(ss, sheetName, formula) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (!formula || typeof formula !== "string" || formula.charAt(0) !== "=") return "ERROR: formula debe empezar con =. Ej: '=SUM(V106:V181)'";

  var targetCell = null;
  try {
    // Calcular EN LA MISMA HOJA para que referencias como =SUM(V106:V181)
    // apunten a la hoja objetivo sin tener que calificar el nombre de hoja.
    var col = sheet.getLastColumn() + 10;
    var row = 1;
    targetCell = sheet.getRange(row, col);
    targetCell.setFormula(formula);
    SpreadsheetApp.flush();
    var result = targetCell.getValue();
    targetCell.clearContent();
    return "Resultado de " + formula + " en " + sheetName + " = " + formatValue(result);
  } catch (e) {
    try { if (targetCell) targetCell.clearContent(); } catch (ignore) {}
    return "ERROR al calcular " + formula + ": " + e.toString();
  }
}

// ═══════════════════════════════════════════════
//  SCHEMA DINÁMICO BASADO EN LA ESTRUCTURA DE BASE
// ═══════════════════════════════════════════════

function buildSchema(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var scanRows = Math.min(lastRow, 30);
  var values = sheet.getRange(1, 1, scanRows, lastCol).getValues();

  var mainHeaderRow = findRowContaining(values, ["obligacion", "capital", "fecha de transferencia"]);
  var mainValueRow = mainHeaderRow ? mainHeaderRow + 1 : 8;
  if (!mainHeaderRow) mainHeaderRow = 7;

  var dailyHeaderRow = findRowContaining(values, ["fecha", "abono realizado a capital"]) || findRowContaining(values, ["fecha", "saldo capital", "causacion"]);
  if (!dailyHeaderRow) dailyHeaderRow = 13;
  var dataStartRow = dailyHeaderRow + 1;

  var fields = mapMainFields(sheet, mainHeaderRow, mainValueRow);
  var daily = mapDailyColumns(sheet, dailyHeaderRow, dataStartRow);

  return {
    hoja: sheet.getName(),
    rows: lastRow,
    cols: lastCol,
    mainHeaderRow: mainHeaderRow,
    mainValueRow: mainValueRow,
    dailyHeaderRow: dailyHeaderRow,
    fields: fields,
    daily: daily
  };
}

function mapMainFields(sheet, headerRow, valueRow) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var result = {}, occurrence = {};
  for (var c = 1; c <= lastCol; c++) {
    var label = String(headers[c - 1] || "").trim();
    if (!label) continue;
    var norm = normalizeKey(label);
    occurrence[norm] = (occurrence[norm] || 0) + 1;
    var key = mainFieldKey(norm, occurrence[norm]);
    if (key) result[key] = { label: label, cell: columnToLetter(c) + valueRow, col: c, row: valueRow };
  }
  return result;
}

function mainFieldKey(norm, occ) {
  if (norm === "obligacion") return "obligacion";
  if (norm === "capital") return "capital";
  if (norm === "fecha de desembolso") return "fecha_desembolso";
  if (norm === "fecha de transferencia") return "fecha_transferencia";
  if (norm === "fecha vcto neto") return "fecha_vcto_neto";
  if (norm === "fecha vcto total") return "fecha_vcto_total";
  if (norm.indexOf("tasa spread proveedor") !== -1 || norm.indexOf("tasa negocio") !== -1) return "tasa_negocio";
  if (norm === "tasa amortizacion") return "tasa_amortizacion";
  if (norm === "dias netos") return "dias_netos";
  if (norm === "valor desembolso") return "valor_desembolso";
  if (norm === "valor descuento") return "valor_descuento";
  if (norm.indexOf("tasa ea interes remuneratorio") !== -1) return "tasa_remuneratorio";
  if (norm === "tasa efectiva diaria") return "tasa_efectiva_diaria";
  if (norm === "dias") return occ === 1 ? "dias_remuneratorio" : "dias_mora";
  if (norm === "valor futuro") return "valor_futuro";
  if (norm === "intereses totales") return "intereses_totales";
  if (norm === "fecha actual") return "fecha_actual";
  if (norm === "saldo capital") return occ === 1 ? "saldo_capital_1" : (occ === 2 ? "saldo_capital_2" : "saldo_capital_3");
  if (norm === "saldo int rem" || norm === "saldo int remuneratorio") return "saldo_int_rem";
  if (norm === "saldo int mora" || norm === "saldo int moratorio") return "saldo_int_mora";
  if (norm === "tasa dpp") return "tasa_dpp";
  if (norm.indexOf("dcto x pronto pago") !== -1) return "dcto_pronto_pago";
  if (norm === "parametro gmf") return "parametro_gmf";
  if (norm === "parametro rte fte") return "parametro_rte_fte";
  if (norm === "valor gmf") return "valor_gmf";
  if (norm === "valor rte fte") return "valor_rte_fte";
  if (norm === "valor giro") return "valor_giro";
  return null;
}

function mapDailyColumns(sheet, headerRow, dataStartRow) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var out = { headerRow: headerRow, dataStartRow: dataStartRow, columns: [] };
  for (var c = 1; c <= lastCol; c++) {
    var label = String(headers[c - 1] || "").trim();
    if (!label) continue;
    var norm = normalizeKey(label);
    out.columns.push({ col: c, letter: columnToLetter(c), label: label });
    if (norm === "fecha") out.dateCol = c;
    if (norm.indexOf("abono realizado a capital") !== -1) out.abonoCapitalCol = c;
    if (norm.indexOf("abono realizado a interes remuneratorio") !== -1) out.abonoRemCol = c;
    if (norm.indexOf("abono realizado a interes moratorio") !== -1) out.abonoMoraCol = c;
    if (norm === "total abono") out.totalAbonoCol = c;
    if (norm === "saldo capital") {
      if (!out.amortSaldoCapitalCol) out.amortSaldoCapitalCol = c;
      else if (!out.remSaldoCapitalCol) out.remSaldoCapitalCol = c;
      else if (!out.moraSaldoCapitalCol) out.moraSaldoCapitalCol = c;
    }
    if (norm.indexOf("intereses remu acumulados") !== -1) out.interesRemAcumCol = c;
    if (norm.indexOf("intereses mora acumulados") !== -1) out.interesMoraAcumCol = c;
    if (norm.indexOf("causacion intres rem") !== -1 || norm.indexOf("causacion interes rem") !== -1) out.interesRemDiarioCol = c;
    if (norm.indexOf("causacion interes mora diario") !== -1) out.interesMoraDiarioCol = c;
    if (norm.indexOf("valor actual de la factura") !== -1 && !out.valorActualCol) out.valorActualCol = c;
  }
  return out;
}

function rowSnapshot(sheet, schema, row) {
  var d = schema.daily;
  var obj = { hoja: sheet.getName(), row: row, fecha: formatValue(sheet.getRange(row, d.dateCol).getValue()) };
  function add(k, col) { if (col) obj[k] = { cell: columnToLetter(col) + row, value: formatValue(sheet.getRange(row, col).getValue()) }; }
  add("saldo_capital", d.amortSaldoCapitalCol || d.remSaldoCapitalCol || d.moraSaldoCapitalCol);
  add("valor_actual_factura", d.valorActualCol);
  add("interes_rem_diario", d.interesRemDiarioCol);
  add("interes_rem_acumulado", d.interesRemAcumCol);
  add("interes_mora_diario", d.interesMoraDiarioCol);
  add("interes_mora_acumulado", d.interesMoraAcumCol);
  add("abono_capital", d.abonoCapitalCol);
  add("abono_interes_rem", d.abonoRemCol);
  add("abono_interes_mora", d.abonoMoraCol);
  add("total_abono", d.totalAbonoCol);
  return obj;
}

// ═══════════════════════════════════════════════
//  PARSEO DE RANGOS A1
// ═══════════════════════════════════════════════

function parseRange(rangeStr) {
  var match = String(rangeStr || "").match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!match) return null;
  var startCol = colLetterToNumber(match[1]);
  var startRow = parseInt(match[2], 10);
  if (match[3]) {
    var endCol = colLetterToNumber(match[3]);
    var endRow = parseInt(match[4], 10);
    return { startCol: startCol, startRow: startRow, endCol: endCol, endRow: endRow, numRows: endRow - startRow + 1, numCols: endCol - startCol + 1 };
  }
  return { startCol: startCol, startRow: startRow, endCol: startCol, endRow: startRow, numRows: 1, numCols: 1 };
}

function colLetterToNumber(letter) {
  var col = 0;
  letter = letter.toUpperCase();
  for (var i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64);
  return col;
}

// ═══════════════════════════════════════════════
//  PROMPT DEL SISTEMA
// ═══════════════════════════════════════════════

function buildSystemPrompt() {
  return ""
    + "Eres un asistente experto en análisis y modificación de un simulador financiero en Google Sheets.\n\n"
    + "CONTEXTO DEL EXCEL:\n"
    + "- Base es la plantilla estructural. Obligaciones lista obligaciones. Tasa usura contiene tasas históricas.\n"
    + "- Las hojas de simulación tienen la misma estructura que Base.\n"
    + "- Base, Obligaciones y Tasa usura son hojas protegidas: no las modifiques salvo que el usuario lo pida explícitamente.\n\n"
    + "USÁ LAS HERRAMIENTAS DISPONIBLES:\n"
    + "- list_sheets: primero para saber qué hojas existen.\n"
    + "- get_simulator_schema: para entender columnas reales por encabezado.\n"
    + "- get_summary: para leer campos principales y última fecha útil.\n"
    + "- find_date_row: para buscar fechas.\n"
    + "- apply_payment: para hacer abonos en una sola llamada.\n"
    + "- clear_payment: para limpiar abonos.\n"
    + "- write_named_field: para cambiar capital, fechas, tasas y parámetros por nombre lógico.\n"
    + "- read_cells/get_cell_info/write_cells: herramientas genéricas si las específicas no alcanzan.\n"
    + "- calculate: PARA TODO CÁLCULO. Nunca calcules manualmente.\n\n"
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
    + "4) Abonos (Y:AB): Abono a Capital, Abono a Int Rem, Abono a Int Mora, Total Abono.\n"
    + "  Columnas de abono se detectan por encabezado, NO por letras fijas. Usá apply_payment.\n\n"
    + "MODIFICACIONES:\n"
    + "- Para un abono: apply_payment. Si el usuario solo da monto, aplica todo a capital; interés remuneratorio=0; interés moratorio=0.\n"
    + "- Si no hay fecha para abono, usa la última fecha útil de la tabla diaria.\n"
    + "- Si no hay monto y el usuario delega en vos, usa 1000000 a capital.\n"
    + "- Para cambiar tasa/capital/fecha: write_named_field.\n"
    + "- Fechas: dd/mm/yyyy. Tasas: decimal (15%=0.15). Montos sin separadores.\n\n"
    + "REGLAS:\n"
    + "- NUNCA inventes valores. Si no leíste un dato con read_cells, no lo asumas.\n"
    + "- NUNCA modifiques Base, Obligaciones o Tasa usura sin permiso explícito.\n"
    + "- Para sumatorias, promedios o CUALQUIER cálculo: usá calculate. NUNCA sumes manualmente.\n"
    + "- Para mostrar tablas usá formato markdown:\n"
    + "  | Fecha | Valor |\n"
    + "  |-------|-------|\n"
    + "  | 10/05 | 208,45 |\n"
    + "- Después de modificar, indicá qué cambiaste, celda(s), fila y valores.\n\n"
    + "ESTILO DE RESPUESTA (OBLIGATORIO):\n"
    + "- Responde en español, conciso y directo al resultado.\n"
    + "- NUNCA uses saludos ni despedidas (Listo, Perfecto, Claro, Entendido).\n"
    + "- NUNCA pidas disculpas ni digas 'disculpe', 'tiene razón'.\n"
    + "- NUNCA repitas la pregunta ni expliques lo obvio.\n"
    + "- Ve DIRECTAMENTE al resultado: tabla, número o dato solicitado.\n"
    + "- Si el usuario dice 'a tu concepto': ACTUÁ sin preguntar. Usá defaults razonables.\n"
    + "- NUNCA le des instrucciones al usuario. VOS sos el asistente, VOS ejecutás.\n"
    + "- Si calculate da 0, es 0. No asumas error.";
}

// ═══════════════════════════════════════════════
//  UTILIDADES
// ═══════════════════════════════════════════════

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function safeJsonParse(text) { try { return JSON.parse(text); } catch (e) { return null; } }
function normalizeText(text) { return String(text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(); }
function normalizeKey(text) { return normalizeText(text).replace(/\./g, "").replace(/\s+/g, " ").trim(); }
function isBlank(v) { return v === "" || v === null || v === undefined; }

function isProtectedSheet(name) {
  var n = normalizeText(name);
  for (var i = 0; i < PROTECTED_SHEETS.length; i++) {
    if (n === PROTECTED_SHEETS[i] || n.indexOf(PROTECTED_SHEETS[i]) !== -1) return true;
  }
  return false;
}

function normalizeInputValue(value, field) {
  if (value === null || value === undefined) return value;
  if (typeof value === "string") {
    var t = value.trim();
    if (t === "") return "";
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(t)) return parseDateFlexible(t);
    if (/^-?[\d\.]+,\d+$/.test(t)) t = t.replace(/\./g, "").replace(",", ".");
    if (/^-?[\d,]+\.\d+$/.test(t)) t = t.replace(/,/g, "");
    if (/^-?\d+(\.\d+)?%$/.test(t)) return parseFloat(t.replace("%", "")) / 100;
    if (/^-?\d+(\.\d+)?$/.test(t)) {
      var num = parseFloat(t);
      if (field && String(field).indexOf("tasa") !== -1 && num > 1) return num / 100;
      return num;
    }
  }
  if (typeof value === "number" && field && String(field).indexOf("tasa") !== -1 && value > 1) return value / 100;
  return value;
}

function toNumberOrZero(v) {
  var n = normalizeInputValue(v);
  if (n === "" || n === null || n === undefined) return 0;
  n = Number(n);
  return isNaN(n) ? 0 : n;
}

function findSheet(ss, name) {
  if (!name) return null;
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var norm = normalizeText(name);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (normalizeText(sheets[i].getName()) === norm) return sheets[i];
  for (var j = 0; j < sheets.length; j++) if (normalizeText(sheets[j].getName()).indexOf(norm) !== -1) return sheets[j];
  return null;
}

function formatValue(val) {
  if (val === "" || val === null || val === undefined) return "vacio";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  if (typeof val === "number") {
    if (Math.abs(val - Math.round(val)) < 0.0000001) return Math.round(val).toString();
    return val.toLocaleString("es-CO", { minimumFractionDigits: 2, maximumFractionDigits: 6 });
  }
  return val.toString();
}

// 1-indexed: col 1 = A, col 26 = Z, col 27 = AA
function columnToLetter(col) {
  var temp = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

function allEmpty(values) {
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (!isBlank(values[r][c])) return false;
    }
  }
  return true;
}

function findRowContaining(values, requiredNorms) {
  for (var r = 0; r < values.length; r++) {
    var rowText = values[r].map(function(v) { return normalizeKey(v); }).join(" | ");
    var ok = true;
    for (var i = 0; i < requiredNorms.length; i++) {
      if (rowText.indexOf(normalizeKey(requiredNorms[i])) === -1) { ok = false; break; }
    }
    if (ok) return r + 1;
  }
  return null;
}

// ── Utilidades de fecha ──

function parseDateFlexible(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") return excelSerialToDate(value);
  var s = String(value || "").trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function coerceToDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000 && value < 80000) return excelSerialToDate(value);
  if (typeof value === "string" && value.trim()) return parseDateFlexible(value);
  return null;
}

function excelSerialToDate(serial) {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

function dateKey(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd"); }

function findRowByDate(sheet, dateCol, startRow, lastRow, dateStr) {
  var target = parseDateFlexible(dateStr);
  if (!target) return null;
  var targetKey = dateKey(target);
  var n = Math.max(0, lastRow - startRow + 1);
  var values = sheet.getRange(startRow, dateCol, n, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var d = coerceToDate(values[i][0]);
    if (d && dateKey(d) === targetKey) return startRow + i;
  }
  return null;
}

function findLastDateRow(sheet, dateCol, startRow, lastRow) {
  var n = Math.max(0, lastRow - startRow + 1);
  var values = sheet.getRange(startRow, dateCol, n, 1).getValues();
  var last = null;
  for (var i = 0; i < values.length; i++) if (coerceToDate(values[i][0])) last = startRow + i;
  return last;
}

function inferHeaderForCell(sheet, schema, row, col) {
  if (row === schema.mainValueRow) {
    for (var k in schema.fields) if (schema.fields[k].col === col) return schema.fields[k].label;
  }
  if (row >= schema.daily.dataStartRow) {
    var label = sheet.getRange(schema.daily.headerRow, col).getValue();
    return label ? String(label) : "";
  }
  var above = sheet.getRange(Math.max(1, row - 1), col).getValue();
  return above ? String(above) : "";
}

function normalizeFieldAlias(field) {
  var k = normalizeKey(field);
  var aliases = {
    "obligacion": "obligacion",
    "capital": "capital",
    "fecha desembolso": "fecha_desembolso",
    "fecha de desembolso": "fecha_desembolso",
    "desembolso": "fecha_desembolso",
    "fecha transferencia": "fecha_transferencia",
    "fecha de transferencia": "fecha_transferencia",
    "fecha vcto neto": "fecha_vcto_neto",
    "vcto neto": "fecha_vcto_neto",
    "fecha vencimiento neto": "fecha_vcto_neto",
    "fecha vcto total": "fecha_vcto_total",
    "vcto total": "fecha_vcto_total",
    "fecha vencimiento total": "fecha_vcto_total",
    "tasa negocio": "tasa_negocio",
    "tasa spread proveedor": "tasa_negocio",
    "tasa proveedor": "tasa_negocio",
    "tasa amortizacion": "tasa_amortizacion",
    "tasa remuneratorio": "tasa_remuneratorio",
    "tasa interes remuneratorio": "tasa_remuneratorio",
    "tasa ea interes remuneratorio": "tasa_remuneratorio",
    "tasa dpp": "tasa_dpp",
    "gmf": "parametro_gmf",
    "parametro gmf": "parametro_gmf",
    "rte fte": "parametro_rte_fte",
    "retefuente": "parametro_rte_fte",
    "parametro rte fte": "parametro_rte_fte"
  };
  return aliases[k] || k;
}
