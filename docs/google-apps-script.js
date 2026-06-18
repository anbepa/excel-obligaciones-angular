/**
 * SIMULADOR FINANCIERO DE OBLIGACIONES — Middleware Google Apps Script + DeepSeek (Function Calling)
 *
 * MODELO REAL DEL EXCEL (hoja "Base" y cada hoja de simulación clonada de Base):
 *  INPUTS EDITABLES (de aquí dependen TODAS las simulaciones):
 *    A8  Obligación        B8  Capital          B4  Fecha de desembolso
 *    C8  Fecha Transferencia  D8  Fecha Vcto Neto   E8  Fecha Vcto Total
 *    F8  Tasa Negocio(EA)  L8  Tasa Int. Remuneratorio(EA)  AB8 Tasa DPP
 *    AD8 Parametro GMF     AE8 Parametro Rte Fte
 *  RESULTADOS (FORMULAS, NO escribir):
 *    I8 Valor Desembolso  J8 Valor Descuento  O8 Valor Futuro  P8 Intereses totales
 *    Q8 Saldo Capital     S8 Fecha Actual(HOY)  Z8 Saldo Int.Rem  AA8 Saldo Int.Mora
 *    AC8 Dcto Pronto Pago AF8 Valor GMF  AG8 Valor Rte Fte  AH8 Valor Giro
 *  CRONOGRAMA DIARIO (eje de fecha unico = columna A, inicia fila 14 = B4, A15=A14+1...):
 *    Fase Amortizacion  (A:J)  filas desembolso -> Vcto Neto
 *    Fase Remuneratorio (L:Q)  filas Vcto Neto  -> Vcto Total
 *    Fase Moratorio     (S:W)  filas Vcto Total -> Fecha Actual
 *    Abonos (Y:AB): Y=Abono Capital  Z=Abono Int.Rem  AA=Abono Int.Mora  AB=Total(FORMULA =SUM(Y:AA))
 *    >> Para abonar SOLO se escriben Y, Z, AA. NUNCA AB (es formula).
 */

var DEEPSEEK_API_KEY = PropertiesService.getScriptProperties().getProperty("DEEPSEEK_API_KEY");
var DEFAULT_FOLDER_ID = "1u51rRx9XiKdzVzBjSg7wYgTOqmuoEjLC";
var MAX_TOOL_TURNS = 24;
var DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions";
var DEEPSEEK_MODEL = "deepseek-chat";
var PROTECTED_SHEETS = ["base", "obligaciones", "tasa usura"];
var MAX_READ_ROWS = 200;

// ===============================================================
//  TOOLS DEFINITION (Function Calling)
// ===============================================================

var TOOLS = [
  {
    type: "function",
    function: {
      name: "list_sheets",
      description: "Lista las hojas del archivo con nombre, filas, columnas y tipo (protegida|simulacion). Base, Obligaciones y Tasa usura son protegidas. Cada hoja de simulacion se llama por el numero de obligacion.",
      parameters: { type: "object", properties: {} }
    }
  },
  {
    type: "function",
    function: {
      name: "get_overview",
      description: "PRIMERA LLAMADA RECOMENDADA. Devuelve TODO el panorama de una hoja en un solo paso: inputs editables (obligacion, capital, fechas, tasas, parametros), resultados calculados (saldos, GMF, Rte Fte, Valor Giro, Dcto Pronto Pago), el cronograma con sus 3 fases y rangos de filas, la fila vigente (Fecha Actual=HOY) y los abonos ya aplicados. Con esto entiendes la obligacion sin leer celda por celda.",
      parameters: {
        type: "object",
        properties: { sheet: { type: "string", description: "Nombre exacto de la hoja de simulacion" } },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_inputs",
      description: "Devuelve solo los 11 campos editables de los que depende la simulacion: A8 Obligacion, B8 Capital, B4 Fecha desembolso, C8 Transferencia, D8 Vcto Neto, E8 Vcto Total, F8 Tasa Negocio, L8 Tasa Remuneratorio, AB8 Tasa DPP, AD8 Parametro GMF, AE8 Parametro Rte Fte.",
      parameters: {
        type: "object",
        properties: { sheet: { type: "string" } },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_simulator_schema",
      description: "Mapa tecnico completo leido de los encabezados reales: celdas de cada campo principal (editable/formula), columnas de cada fase del cronograma y columnas de abonos. Usalo si necesitas las letras de columna exactas.",
      parameters: {
        type: "object",
        properties: { sheet: { type: "string" } },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_schedule",
      description: "Devuelve filas del cronograma diario con columnas con nombre (fecha, saldo_capital, causacion_rem_dia, int_rem_acum, causacion_mora_dia, int_mora_acum, abono_total). Sin fechas: ventana alrededor de la Fecha Actual. Con from/to: ese rango. Ideal para mostrar tablas.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          from: { type: "string", description: "Fecha inicio dd/mm/yyyy. Opcional." },
          to: { type: "string", description: "Fecha fin dd/mm/yyyy. Opcional." },
          maxRows: { type: "number", description: "Maximo de filas a devolver. Default 30." }
        },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "find_date_row",
      description: "Busca una fecha dd/mm/yyyy en el eje del cronograma (columna A) y devuelve el estado de esa fila (saldos, causaciones, abonos). Sin fecha devuelve la fila vigente (Fecha Actual = HOY).",
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
      description: "Aplica un abono escribiendo SOLO las columnas Y (capital), Z (int. remuneratorio) y AA (int. moratorio). El Total (AB) es formula y se recalcula solo. Encuentra la fila por fecha; si no se da fecha usa la fila vigente (HOY). Devuelve el impacto en saldos y Valor Giro.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          date: { type: "string", description: "Fecha dd/mm/yyyy. Opcional; default fila vigente (HOY)." },
          capital: { type: "number", description: "Abono a capital. Si el usuario no discrimina, todo el monto va aqui." },
          interestRem: { type: "number", description: "Abono a interes remuneratorio. Default 0." },
          interestMora: { type: "number", description: "Abono a interes moratorio. Default 0." }
        },
        required: ["sheet"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "clear_payment",
      description: "Limpia el abono de una fecha borrando SOLO Y, Z y AA (nunca AB, que es formula). Sin fecha usa la fila vigente (HOY).",
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
      name: "write_named_field",
      description: "Modifica UN input editable por nombre logico. Permitidos: obligacion, capital, fecha_desembolso, fecha_transferencia, fecha_vcto_neto, fecha_vcto_total, tasa_negocio, tasa_remuneratorio, tasa_dpp, parametro_gmf, parametro_rte_fte. Rechaza campos calculados por formula. Devuelve el valor escrito y el recalculo de saldos.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          field: { type: "string" },
          value: { description: "Valor. Fechas dd/mm/yyyy; tasas en % o decimal; montos como numero." }
        },
        required: ["sheet", "field", "value"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "read_cells",
      description: "Lee valores reales de un rango A1. Maximo 200 filas. Ej: 'B8', 'A14:J50', 'Y8:AH8'. Muestra valor y formula entre llaves cuando existe.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string", description: "Rango en notacion A1" }
        },
        required: ["sheet", "range"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "get_cell_info",
      description: "Valor, formula, nota, si es editable o calculada, y encabezado probable de una celda.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          cell: { type: "string", description: "Celda A1. Ej: 'B8', 'AH8', 'Q79'" }
        },
        required: ["sheet", "cell"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "write_cells",
      description: "Escribe valores en un rango A1. Generico: usalo solo si no hay tool especifica. Por seguridad NO sobrescribe celdas con formula salvo force=true. Para limpiar usa valores vacios ''.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          range: { type: "string" },
          values: { type: "array", items: { type: "array", items: {} } },
          force: { type: "boolean", description: "true para permitir sobrescribir formulas. Default false." }
        },
        required: ["sheet", "range", "values"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "calculate",
      description: "Ejecuta una formula en la hoja objetivo y devuelve el resultado. USALA SIEMPRE para sumas, promedios, restas, conteos. Referencias sin calificar apuntan a la hoja indicada. Ej: '=SUM(V68:V79)'.",
      parameters: {
        type: "object",
        properties: {
          sheet: { type: "string" },
          formula: { type: "string", description: "Formula completa con =." }
        },
        required: ["sheet", "formula"]
      }
    }
  },
  {
    type: "function",
    function: {
      name: "read_usura_rate",
      description: "Consulta la tasa de usura historica en la hoja 'Tasa usura'. Con fecha dd/mm/yyyy devuelve la tasa de ese dia; sin fecha devuelve la ultima disponible.",
      parameters: {
        type: "object",
        properties: { date: { type: "string", description: "Fecha dd/mm/yyyy. Opcional." } },
        required: []
      }
    }
  }
];

// ===============================================================
//  ENDPOINTS HTTP
// ===============================================================

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
    return jsonResponse({ success: false, error: "Accion no valida" });
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ===============================================================
//  CHAT CON FUNCTION CALLING
// ===============================================================

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

      if (!result) return jsonResponse({ success: false, error: "DeepSeek no devolvio JSON valido: " + raw.substring(0, 500) });
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
    if (!finalReply) return jsonResponse({ success: false, error: "El asistente no genero respuesta despues de " + turn + " turnos. Intenta una pregunta mas concreta." });

    return jsonResponse({ success: true, reply: finalReply });
  } catch (error) {
    return jsonResponse({ success: false, error: error.toString() });
  }
}

// ===============================================================
//  EJECUCION DE TOOLS
// ===============================================================

function executeTool(ss, toolName, args) {
  try {
    if (toolName === "list_sheets") return toolListSheets(ss);
    if (toolName === "get_overview") return toolGetOverview(ss, args.sheet);
    if (toolName === "get_inputs") return toolGetInputs(ss, args.sheet);
    if (toolName === "get_simulator_schema") return toolGetSimulatorSchema(ss, args.sheet);
    if (toolName === "get_schedule") return toolGetSchedule(ss, args.sheet, args.from, args.to, args.maxRows);
    if (toolName === "find_date_row") return toolFindDateRow(ss, args.sheet, args.date);
    if (toolName === "apply_payment") return toolApplyPayment(ss, args.sheet, args.date, args.capital, args.interestRem, args.interestMora);
    if (toolName === "clear_payment") return toolClearPayment(ss, args.sheet, args.date);
    if (toolName === "write_named_field") return toolWriteNamedField(ss, args.sheet, args.field, args.value);
    if (toolName === "get_cell_info") return toolGetCellInfo(ss, args.sheet, args.cell);
    if (toolName === "read_cells") return toolReadCells(ss, args.sheet, args.range);
    if (toolName === "write_cells") return toolWriteCells(ss, args.sheet, args.range, args.values, args.force);
    if (toolName === "calculate") return toolCalculate(ss, args.sheet, args.formula);
    if (toolName === "read_usura_rate") return toolReadUsura(ss, args.date);
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

function toolGetOverview(ss, sheetName) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  var d = schema.daily;

  var inputs = {}, resultados = {};
  var keys = Object.keys(schema.fields);
  for (var i = 0; i < keys.length; i++) {
    var meta = schema.fields[keys[i]];
    var entry = { cell: meta.cell, value: formatValue(sheet.getRange(meta.cell).getValue()) };
    if (meta.editable) inputs[keys[i]] = entry; else resultados[keys[i]] = entry;
  }

  var currentRow = findCurrentRow(sheet, schema);
  var fechaActual = schema.fields.fecha_actual ? formatValue(sheet.getRange(schema.fields.fecha_actual.cell).getValue()) : null;

  var fases = {};
  if (d.amortSaldoCapCol) { var a = colSpanNumeric(sheet, d.amortSaldoCapCol, d.dataStartRow); if (a.first) fases.amortizacion = { cols: "A:J", filas: a.first + ".." + a.last }; }
  if (d.remSaldoCapCol) { var r = colSpanNumeric(sheet, d.remSaldoCapCol, d.dataStartRow); if (r.first) fases.remuneratorio = { cols: "L:Q", filas: r.first + ".." + r.last }; }
  if (d.moraSaldoCapCol) { var m = colSpanNumeric(sheet, d.moraSaldoCapCol, d.dataStartRow); if (m.first) fases.moratorio = { cols: "S:W", filas: m.first + ".." + m.last }; }

  var abonos = scanAbonos(sheet, schema);

  var out = {
    hoja: sheet.getName(),
    tipo: isProtectedSheet(sheet.getName()) ? "protegida" : "simulacion",
    inputs_editables: inputs,
    resultados_calculados: resultados,
    cronograma: {
      eje_fecha_col: "A",
      fila_inicio: d.dataStartRow,
      fecha_actual: fechaActual,
      fila_vigente: currentRow,
      fases: fases,
      abonos_cols: { capital: "Y", int_rem: "Z", int_mora: "AA", total: "AB (formula =SUM(Y:AA))" }
    },
    abonos_aplicados: abonos.length ? abonos : "ninguno"
  };
  return JSON.stringify(out);
}

function toolGetInputs(ss, sheetName) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  var order = ["obligacion","capital","fecha_desembolso","fecha_transferencia","fecha_vcto_neto","fecha_vcto_total","tasa_negocio","tasa_remuneratorio","tasa_dpp","parametro_gmf","parametro_rte_fte"];
  var out = { hoja: sheet.getName(), inputs: {} };
  for (var i = 0; i < order.length; i++) {
    var meta = schema.fields[order[i]];
    if (meta) out.inputs[order[i]] = { label: meta.label, cell: meta.cell, value: formatValue(sheet.getRange(meta.cell).getValue()) };
  }
  return JSON.stringify(out);
}

function toolGetSimulatorSchema(ss, sheetName) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  return JSON.stringify(buildSchema(sheet));
}

function toolGetSchedule(ss, sheetName, fromStr, toStr, maxRows) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  var d = schema.daily;
  if (!d.dateCol) return "ERROR: no encontre la columna Fecha del cronograma.";
  var lastSheetRow = sheet.getLastRow();
  var startRow, endRow;
  var cap = maxRows && maxRows > 0 ? Math.min(maxRows, MAX_READ_ROWS) : 30;

  if (fromStr || toStr) {
    startRow = fromStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, lastSheetRow, fromStr) : d.dataStartRow;
    endRow = toStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, lastSheetRow, toStr) : (startRow + cap - 1);
    if (!startRow) return "ERROR: fecha 'from' no encontrada: " + fromStr;
    if (!endRow) endRow = startRow + cap - 1;
  } else {
    var cur = findCurrentRow(sheet, schema) || findLastActiveRow(sheet, schema) || d.dataStartRow;
    startRow = Math.max(d.dataStartRow, cur - cap + 1);
    endRow = cur;
  }
  if (endRow < startRow) { var tmp = startRow; startRow = endRow; endRow = tmp; }
  if (endRow - startRow + 1 > cap) startRow = endRow - cap + 1;

  // Batch: leer todas las columnas necesarias en un solo getValues()
  var colsToFetch = [d.dateCol];
  var saldoCols = [d.amortSaldoCapCol, d.remSaldoCapCol, d.moraSaldoCapCol].filter(Boolean);
  saldoCols.forEach(function(c) { colsToFetch.push(c); });
  if (d.remCausCol) colsToFetch.push(d.remCausCol);
  if (d.remAcumCol) colsToFetch.push(d.remAcumCol);
  if (d.moraCausCol) colsToFetch.push(d.moraCausCol);
  if (d.moraAcumCol) colsToFetch.push(d.moraAcumCol);
  if (d.totalAbonoCol) colsToFetch.push(d.totalAbonoCol);

  var minC = Math.min.apply(null, colsToFetch);
  var maxC = Math.max.apply(null, colsToFetch);
  var block = sheet.getRange(startRow, minC, endRow - startRow + 1, maxC - minC + 1).getValues();

  var colIdx = {};
  colsToFetch.forEach(function(c) { colIdx[c] = c - minC; });

  var rows = [];
  for (var rr = 0; rr < block.length; rr++) {
    var rowData = block[rr];
    var absRow = startRow + rr;
    var rec = { fila: absRow, fecha: formatValue(rowData[colIdx[d.dateCol]]) };
    rec.saldo_capital = pickFromRow(rowData, colIdx, saldoCols);
    if (d.remCausCol) rec.causacion_rem_dia = formatValue(rowData[colIdx[d.remCausCol]]);
    if (d.remAcumCol) rec.int_rem_acum = formatValue(rowData[colIdx[d.remAcumCol]]);
    if (d.moraCausCol) rec.causacion_mora_dia = formatValue(rowData[colIdx[d.moraCausCol]]);
    if (d.moraAcumCol) rec.int_mora_acum = formatValue(rowData[colIdx[d.moraAcumCol]]);
    if (d.totalAbonoCol) rec.abono_total = formatValue(rowData[colIdx[d.totalAbonoCol]]);
    rows.push(rec);
  }
  return JSON.stringify({ hoja: sheet.getName(), desde: startRow, hasta: endRow, filas: rows });
}

function toolFindDateRow(ss, sheetName, dateStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var schema = buildSchema(sheet);
  if (!schema.daily.dateCol) return "ERROR: no encontre columna Fecha en el cronograma.";
  var row = dateStr ? findRowByDate(sheet, schema.daily.dateCol, schema.daily.dataStartRow, sheet.getLastRow(), dateStr) : findCurrentRow(sheet, schema);
  if (!row) return "ERROR: fecha no encontrada: " + (dateStr || "fila vigente (HOY)");
  return JSON.stringify(rowSnapshot(sheet, schema, row));
}

function toolApplyPayment(ss, sheetName, dateStr, capital, interestRem, interestMora) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura). Usa una hoja de simulacion.";

  var schema = buildSchema(sheet);
  var d = schema.daily;
  if (!d.dateCol || !d.abonoCapitalCol || !d.abonoRemCol || !d.abonoMoraCol) {
    return "ERROR: no encontre columnas de Abonos (Y/Z/AA). Ejecuta get_simulator_schema.";
  }

  var row = dateStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow(), dateStr) : findCurrentRow(sheet, schema);
  if (!row) return "ERROR: fecha de abono no encontrada: " + (dateStr || "fila vigente (HOY)");

  var cap = toNumberOrZero(capital);
  var rem = toNumberOrZero(interestRem);
  var mora = toNumberOrZero(interestMora);

  // Escribe SOLO Y (capital), Z (rem), AA (mora). AB (Total) es formula =SUM(Y:AA) y se recalcula solo.
  var startCol = Math.min(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol);
  var endCol = Math.max(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol);
  var width = endCol - startCol + 1;
  var vals = sheet.getRange(row, startCol, 1, width).getValues()[0];
  vals[d.abonoCapitalCol - startCol] = cap;
  vals[d.abonoRemCol - startCol] = rem;
  vals[d.abonoMoraCol - startCol] = mora;
  sheet.getRange(row, startCol, 1, width).setValues([vals]);
  SpreadsheetApp.flush();

  var impacto = readOutputs(sheet, schema);
  var totalCell = d.totalAbonoCol ? columnToLetter(d.totalAbonoCol) + row : null;

  return JSON.stringify({
    ok: true,
    accion: "abono_aplicado",
    hoja: sheet.getName(),
    fila: row,
    fecha: formatValue(sheet.getRange(row, d.dateCol).getValue()),
    celdas_escritas: {
      capital: columnToLetter(d.abonoCapitalCol) + row,
      interes_remuneratorio: columnToLetter(d.abonoRemCol) + row,
      interes_moratorio: columnToLetter(d.abonoMoraCol) + row
    },
    valores: { capital: cap, interes_remuneratorio: rem, interes_moratorio: mora },
    total_abono_celda: totalCell,
    total_abono: totalCell ? formatValue(sheet.getRange(totalCell).getValue()) : null,
    impacto_recalculado: impacto
  });
}

function toolClearPayment(ss, sheetName, dateStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var schema = buildSchema(sheet);
  var d = schema.daily;
  if (!d.dateCol || !d.abonoCapitalCol || !d.abonoRemCol || !d.abonoMoraCol) return "ERROR: no encontre columnas de Abonos.";
  var row = dateStr ? findRowByDate(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow(), dateStr) : findCurrentRow(sheet, schema);
  if (!row) return "ERROR: fecha no encontrada: " + (dateStr || "fila vigente (HOY)");
  // Limpia SOLO Y, Z, AA (nunca AB, que es formula).
  var startCol = Math.min(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol);
  var endCol = Math.max(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol);
  sheet.getRange(row, startCol, 1, endCol - startCol + 1).clearContent();
  SpreadsheetApp.flush();
  return "Limpiado abono en " + sheet.getName() + " fila " + row + " (" + columnToLetter(startCol) + row + ":" + columnToLetter(endCol) + row + ", AB intacta).";
}

function toolWriteNamedField(ss, sheetName, field, value) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var schema = buildSchema(sheet);
  var key = normalizeFieldAlias(field);
  var meta = schema.fields[key];
  if (!meta) {
    var editables = [];
    for (var k in schema.fields) if (schema.fields[k].editable) editables.push(k);
    return "ERROR: campo no encontrado: " + field + ". Editables: " + editables.join(", ");
  }
  if (!meta.editable) return "ERROR: '" + key + "' (" + meta.cell + ") es un valor CALCULADO por formula, no se puede editar. Editables: A8,B8,B4,C8,D8,E8,F8,L8,AB8,AD8,AE8.";
  var v = normalizeInputValue(value, meta.type === "rate" ? "tasa" : key);
  sheet.getRange(meta.cell).setValue(v);
  SpreadsheetApp.flush();
  return JSON.stringify({
    ok: true,
    hoja: sheet.getName(),
    campo: key,
    label: meta.label,
    celda: meta.cell,
    valor: formatValue(sheet.getRange(meta.cell).getValue()),
    impacto_recalculado: readOutputs(sheet, schema)
  });
}

function toolGetCellInfo(ss, sheetName, cellRef) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var parsed = parseRange(cellRef);
  if (!parsed || parsed.numRows !== 1 || parsed.numCols !== 1) return "ERROR: '" + cellRef + "' no es una celda valida. Usa formato A1.";
  var range = sheet.getRange(parsed.startRow, parsed.startCol);
  var value = range.getValue();
  var formula = range.getFormula();
  var note = "";
  try { note = range.getNote() || ""; } catch (e) {}
  var schema = buildSchema(sheet);
  var header = inferHeaderForCell(sheet, schema, parsed.startRow, parsed.startCol);
  var lines = ["Celda: " + cellRef, "Encabezado probable: " + (header || "vacio"), "Tipo: " + (formula ? "CALCULADA (formula)" : "valor editable")];
  if (formula) { lines.push("Formula: " + formula); lines.push("Valor calculado: " + formatValue(value)); }
  else { lines.push("Valor: " + formatValue(value)); }
  if (note) lines.push("Nota: " + note);
  return lines.join("\n");
}

function toolReadCells(ss, sheetName, rangeStr) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango invalido '" + rangeStr + "'. Usa notacion A1.";
  var numRows = parsed.numRows;
  var numCols = parsed.numCols;
  if (numRows > MAX_READ_ROWS) { numRows = MAX_READ_ROWS; parsed.endRow = parsed.startRow + MAX_READ_ROWS - 1; }
  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);
  var values = range.getValues();
  var formulas = range.getFormulas();
  if (numRows === 1 && numCols === 1) {
    var v = values[0][0], f = formulas[0][0];
    return f ? formatValue(v) + " (formula: " + f + ")" : formatValue(v);
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
  if (rows.length === 0) return "(rango vacio)";
  return rows.join("\n") + (numRows < parsed.numRows ? "\n[Truncado a 200 filas]" : "");
}

function toolWriteCells(ss, sheetName, rangeStr, values, force) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (isProtectedSheet(sheet.getName())) return "ERROR: no modifico hojas protegidas (Base, Obligaciones, Tasa usura).";
  var parsed = parseRange(rangeStr);
  if (!parsed) return "ERROR: rango invalido '" + rangeStr + "'.";
  if (!values || !values.length || !values[0].length) return "ERROR: values vacio o invalido.";

  var numRows = Math.min(values.length, parsed.numRows);
  var numCols = Math.min(values[0].length, parsed.numCols);
  var range = sheet.getRange(parsed.startRow, parsed.startCol, numRows, numCols);

  if (!force) {
    var existing = range.getFormulas();
    var blocked = [];
    for (var fr = 0; fr < existing.length; fr++) {
      for (var fc = 0; fc < existing[fr].length; fc++) {
        if (existing[fr][fc]) blocked.push(columnToLetter(parsed.startCol + fc) + (parsed.startRow + fr));
      }
    }
    if (blocked.length) return "ERROR: el rango contiene formulas (" + blocked.slice(0, 8).join(", ") + (blocked.length > 8 ? "..." : "") + "). No las sobrescribo. Si es intencional, repite con force=true.";
  }

  var writeValues = [];
  for (var rr = 0; rr < numRows; rr++) {
    var rowOut = [];
    for (var cc = 0; cc < numCols; cc++) rowOut.push(normalizeInputValue(values[rr][cc]));
    writeValues.push(rowOut);
  }
  if (allEmpty(writeValues)) { range.clearContent(); SpreadsheetApp.flush(); return "Limpiado " + sheetName + "!" + rangeStr; }
  range.setValues(writeValues);
  SpreadsheetApp.flush();
  return "Escrito en " + sheetName + "!" + rangeStr + " -> " + JSON.stringify(writeValues).substring(0, 200);
}

function toolCalculate(ss, sheetName, formula) {
  var sheet = findSheet(ss, sheetName);
  if (!sheet) return "ERROR: hoja '" + sheetName + "' no encontrada.";
  if (!formula || typeof formula !== "string" || formula.charAt(0) !== "=") return "ERROR: formula debe empezar con =. Ej: '=SUM(V68:V79)'";

  var targetCell = null;
  try {
    // Calcula en una celda auxiliar lejos de los datos, en la MISMA hoja,
    // para que referencias como =SUM(V68:V79) apunten a la hoja objetivo.
    var col = Math.min(sheet.getMaxColumns(), sheet.getLastColumn() + 50);
    targetCell = sheet.getRange(1, col);
    targetCell.setFormula(formula);
    SpreadsheetApp.flush();
    var resultValue = targetCell.getValue();
    targetCell.clearContent();
    return "Resultado de " + formula + " en " + sheetName + " = " + formatValue(resultValue);
  } catch (e) {
    try { if (targetCell) targetCell.clearContent(); } catch (ignore) {}
    return "ERROR al calcular " + formula + ": " + e.toString();
  }
}

function toolReadUsura(ss, dateStr) {
  var sheet = ss.getSheetByName("Tasa usura") || findSheet(ss, "tasa usura");
  if (!sheet) return "ERROR: no existe la hoja 'Tasa usura'.";
  var last = sheet.getLastRow();
  // Localiza fila de encabezado 'Fecha' en col A (suele ser fila 2) y arranca debajo.
  var headerRow = 2;
  var topA = sheet.getRange(1, 1, Math.min(5, last), 1).getValues();
  for (var i = 0; i < topA.length; i++) { if (normalizeKey(topA[i][0]) === "fecha") { headerRow = i + 1; break; } }
  var start = headerRow + 1;
  var n = Math.max(0, last - start + 1);
  var data = sheet.getRange(start, 1, n, 2).getValues();
  if (dateStr) {
    var target = parseDateFlexible(dateStr);
    if (!target) return "ERROR: fecha invalida: " + dateStr;
    var tkey = dateKey(target);
    for (var j = 0; j < data.length; j++) {
      var dd = coerceToDate(data[j][0]);
      if (dd && dateKey(dd) === tkey) return JSON.stringify({ fecha: formatValue(dd), tasa_usura: formatValue(data[j][1]) });
    }
    return "ERROR: no hay tasa de usura para " + dateStr + ".";
  }
  for (var k = data.length - 1; k >= 0; k--) {
    var d2 = coerceToDate(data[k][0]);
    if (d2 && !isBlank(data[k][1])) return JSON.stringify({ fecha: formatValue(d2), tasa_usura: formatValue(data[k][1]), nota: "ultima disponible" });
  }
  return "ERROR: no encontre tasas de usura.";
}

// ===============================================================
//  SCHEMA DINAMICO BASADO EN LA ESTRUCTURA DE BASE
// ===============================================================

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
  addDisbursementField(sheet, fields, values);
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

// Localiza "Fecha de desembolso" (suele estar en A4 con su valor en B4).
function addDisbursementField(sheet, fields, values) {
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (normalizeKey(values[r][c]) === "fecha de desembolso") {
        var valCol = c + 2; // celda a la derecha de la etiqueta
        fields.fecha_desembolso = {
          label: "Fecha de desembolso",
          cell: columnToLetter(valCol) + (r + 1),
          col: valCol, row: r + 1, editable: true, type: "date"
        };
        return;
      }
    }
  }
  if (!fields.fecha_desembolso) {
    fields.fecha_desembolso = { label: "Fecha de desembolso", cell: "B4", col: 2, row: 4, editable: true, type: "date" };
  }
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
    var meta = classifyMainField(norm, occurrence[norm]);
    if (meta) result[meta.key] = { label: label, cell: columnToLetter(c) + valueRow, col: c, row: valueRow, editable: meta.editable, type: meta.type };
  }
  return result;
}

// Devuelve {key, editable, type} o null. editable=input de la simulacion; !editable=formula.
function classifyMainField(norm, occ) {
  if (norm === "obligacion") return { key: "obligacion", editable: true, type: "text" };
  if (norm === "capital") return { key: "capital", editable: true, type: "number" };
  if (norm === "fecha de transferencia") return { key: "fecha_transferencia", editable: true, type: "date" };
  if (norm === "fecha vcto neto") return { key: "fecha_vcto_neto", editable: true, type: "date" };
  if (norm === "fecha vcto total") return { key: "fecha_vcto_total", editable: true, type: "date" };
  if (norm.indexOf("tasa spread proveedor") !== -1 || norm.indexOf("tasa negocio") !== -1) return { key: "tasa_negocio", editable: true, type: "rate" };
  if (norm.indexOf("tasa ea interes remuneratorio") !== -1) return { key: "tasa_remuneratorio", editable: true, type: "rate" };
  if (norm === "tasa dpp") return { key: "tasa_dpp", editable: true, type: "rate" };
  if (norm === "parametro gmf") return { key: "parametro_gmf", editable: true, type: "number" };
  if (norm === "parametro rte fte") return { key: "parametro_rte_fte", editable: true, type: "number" };
  // Calculados (formula):
  if (norm === "tasa amortizacion") return { key: "tasa_amortizacion", editable: false, type: "rate" };
  if (norm === "dias netos") return { key: "dias_netos", editable: false, type: "number" };
  if (norm === "valor desembolso") return { key: "valor_desembolso", editable: false, type: "number" };
  if (norm === "valor descuento") return { key: "valor_descuento", editable: false, type: "number" };
  if (norm === "tasa efectiva diaria") return { key: "tasa_efectiva_diaria", editable: false, type: "rate" };
  if (norm === "valor futuro") return { key: "valor_futuro", editable: false, type: "number" };
  if (norm === "intereses totales") return { key: "intereses_totales", editable: false, type: "number" };
  if (norm === "fecha actual") return { key: "fecha_actual", editable: false, type: "date" };
  if (norm === "dias") return occ === 1 ? { key: "dias_remuneratorio", editable: false, type: "number" } : { key: "dias_mora", editable: false, type: "number" };
  if (norm === "saldo capital") {
    if (occ === 1) return { key: "saldo_capital", editable: false, type: "number" };
    if (occ === 2) return { key: "saldo_capital_neto", editable: false, type: "number" };
    return { key: "saldo_capital_abonos", editable: false, type: "number" };
  }
  if (norm === "saldo int rem" || norm === "saldo int remuneratorio") return { key: "saldo_int_rem", editable: false, type: "number" };
  if (norm === "saldo int mora" || norm === "saldo int moratorio") return { key: "saldo_int_mora", editable: false, type: "number" };
  if (norm.indexOf("dcto x pronto pago") !== -1) return { key: "dcto_pronto_pago", editable: false, type: "number" };
  if (norm === "valor gmf") return { key: "valor_gmf", editable: false, type: "number" };
  if (norm === "valor rte fte") return { key: "valor_rte_fte", editable: false, type: "number" };
  if (norm === "valor giro") return { key: "valor_giro", editable: false, type: "number" };
  return null;
}

function mapDailyColumns(sheet, headerRow, dataStartRow) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var out = { headerRow: headerRow, dataStartRow: dataStartRow, columns: [] };
  var occ = {};
  for (var c = 1; c <= lastCol; c++) {
    var label = String(headers[c - 1] || "").trim();
    if (!label) continue;
    var norm = normalizeKey(label);
    occ[norm] = (occ[norm] || 0) + 1;
    out.columns.push({ col: c, letter: columnToLetter(c), label: label });

    if (norm === "fecha" && !out.dateCol) out.dateCol = c;
    if (norm.indexOf("dias faltante") !== -1) out.diasFaltCol = c;
    if (norm === "saldo capital") {
      if (occ[norm] === 1) out.amortSaldoCapCol = c;
      else if (occ[norm] === 2) out.remSaldoCapCol = c;
      else if (occ[norm] === 3) out.moraSaldoCapCol = c;
    }
    if (norm.indexOf("valor actual de la factura") !== -1 && !out.amortValorActualCol) out.amortValorActualCol = c;
    if (norm.indexOf("amortizacion diaria") !== -1 && !out.amortDiariaCol) out.amortDiariaCol = c;
    if (norm.indexOf("tasa int remuneratorio") !== -1) out.remTasaCol = c;
    if (norm === "dias") { if (occ[norm] === 1) out.remDiasCol = c; else out.moraDiasCol = c; }
    if (norm.indexOf("valor futuro") !== -1 && !out.remValorFuturoCol) out.remValorFuturoCol = c;
    if ((norm.indexOf("causacion intres rem") !== -1 || norm.indexOf("causacion interes rem") !== -1) && !out.remCausCol) out.remCausCol = c;
    if (norm.indexOf("intereses remu acumulados") !== -1) out.remAcumCol = c;
    if (norm.indexOf("tasa int mora") !== -1) out.moraTasaCol = c;
    if (norm.indexOf("causacion interes mora") !== -1 && !out.moraCausCol) out.moraCausCol = c;
    if (norm.indexOf("intereses mora acumulados") !== -1) out.moraAcumCol = c;
    if (norm.indexOf("abono realizado a capital") !== -1) out.abonoCapitalCol = c;
    if (norm.indexOf("abono realizado a interes remuneratorio") !== -1) out.abonoRemCol = c;
    if (norm.indexOf("abono realizado a interes moratorio") !== -1) out.abonoMoraCol = c;
    if (norm === "total abono") out.totalAbonoCol = c;
  }
  return out;
}

function rowSnapshot(sheet, schema, row) {
  var d = schema.daily;
  var obj = { hoja: sheet.getName(), row: row, fecha: formatValue(sheet.getRange(row, d.dateCol).getValue()) };
  function add(k, col) { if (col) obj[k] = { cell: columnToLetter(col) + row, value: formatValue(sheet.getRange(row, col).getValue()) }; }
  add("saldo_capital", d.amortSaldoCapCol || d.remSaldoCapCol || d.moraSaldoCapCol);
  add("valor_actual_factura", d.amortValorActualCol);
  add("causacion_rem_diaria", d.remCausCol);
  add("int_rem_acumulado", d.remAcumCol);
  add("causacion_mora_diaria", d.moraCausCol);
  add("int_mora_acumulado", d.moraAcumCol);
  add("abono_capital", d.abonoCapitalCol);
  add("abono_interes_rem", d.abonoRemCol);
  add("abono_interes_mora", d.abonoMoraCol);
  add("total_abono", d.totalAbonoCol);
  return obj;
}

// Lee los principales outputs calculados de la fila 8 (resumen del impacto).
function readOutputs(sheet, schema) {
  var f = schema.fields;
  var out = {};
  function add(k) { if (f[k]) out[k] = formatValue(sheet.getRange(f[k].cell).getValue()); }
  add("saldo_capital"); add("saldo_int_rem"); add("saldo_int_mora");
  add("intereses_totales"); add("dcto_pronto_pago");
  add("valor_gmf"); add("valor_rte_fte"); add("valor_giro");
  return out;
}

function scanAbonos(sheet, schema) {
  var d = schema.daily;
  if (!d.abonoCapitalCol || !d.dateCol) return [];
  var last = sheet.getLastRow();
  var start = d.dataStartRow;
  var n = Math.max(0, last - start + 1);
  var startCol = Math.min(d.abonoCapitalCol, d.abonoRemCol, d.abonoMoraCol);
  var endCol = d.totalAbonoCol ? Math.max(d.abonoMoraCol, d.totalAbonoCol) : Math.max(d.abonoRemCol, d.abonoMoraCol);
  var width = endCol - startCol + 1;
  var dates = sheet.getRange(start, d.dateCol, n, 1).getValues();
  var block = sheet.getRange(start, startCol, n, width).getValues();
  var res = [];
  for (var i = 0; i < n; i++) {
    var cap = toNum(block[i][d.abonoCapitalCol - startCol]);
    var rem = toNum(block[i][d.abonoRemCol - startCol]);
    var mora = toNum(block[i][d.abonoMoraCol - startCol]);
    if (cap || rem || mora) {
      res.push({
        fila: start + i,
        fecha: formatValue(dates[i][0]),
        capital: cap, interes_rem: rem, interes_mora: mora,
        total: d.totalAbonoCol ? toNum(block[i][d.totalAbonoCol - startCol]) : (cap + rem + mora)
      });
    }
  }
  return res;
}

// ===============================================================
//  LOCALIZACION DE FILAS DEL CRONOGRAMA
// ===============================================================

function findCurrentRow(sheet, schema) {
  var d = schema.daily;
  var faCell = schema.fields.fecha_actual ? schema.fields.fecha_actual.cell : "S8";
  var fa = coerceToDate(sheet.getRange(faCell).getValue());
  if (fa) {
    var r = findRowByDateObj(sheet, d.dateCol, d.dataStartRow, sheet.getLastRow(), fa);
    if (r) return r;
  }
  return findLastActiveRow(sheet, schema);
}

function findLastActiveRow(sheet, schema) {
  var d = schema.daily;
  var start = d.dataStartRow;
  var last = sheet.getLastRow();
  var cols = [];
  if (d.amortSaldoCapCol) cols.push(d.amortSaldoCapCol);
  if (d.remSaldoCapCol) cols.push(d.remSaldoCapCol);
  if (d.moraSaldoCapCol) cols.push(d.moraSaldoCapCol);
  if (!cols.length) return null;
  var minC = Math.min.apply(null, cols), maxC = Math.max.apply(null, cols);
  var n = Math.max(0, last - start + 1);
  var rng = sheet.getRange(start, minC, n, maxC - minC + 1).getValues();
  var lastRow = null;
  for (var i = 0; i < n; i++) {
    for (var k = 0; k < cols.length; k++) {
      var v = rng[i][cols[k] - minC];
      if (typeof v === "number" && !isNaN(v)) { lastRow = start + i; break; }
    }
  }
  return lastRow;
}

function colSpanNumeric(sheet, col, start) {
  var last = sheet.getLastRow();
  var n = Math.max(0, last - start + 1);
  var vals = sheet.getRange(start, col, n, 1).getValues();
  var first = null, lastR = null;
  for (var i = 0; i < n; i++) {
    var v = vals[i][0];
    if (typeof v === "number" && !isNaN(v)) { if (first === null) first = start + i; lastR = start + i; }
  }
  return { first: first, last: lastR };
}

// ===============================================================
//  PARSEO DE RANGOS A1
// ===============================================================

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

// ===============================================================
//  PROMPT DEL SISTEMA
// ===============================================================

function buildSystemPrompt() {
  return ""
    + "Eres un asistente experto en un SIMULADOR FINANCIERO DE OBLIGACIONES en Google Sheets. Operas con autonomia total: lees, calculas y modificas el sheet con las herramientas; el usuario no interviene.\n\n"
    + "QUE ES CADA HOJA:\n"
    + "- 'Base' es la plantilla estructural. 'Tasa usura' tiene tasas historicas. Cada hoja de simulacion se llama por el numero de obligacion y es un clon de Base con distintos inputs.\n"
    + "- Base y Tasa usura son PROTEGIDAS: no las modifiques salvo orden explicita. Trabaja sobre la hoja de simulacion.\n\n"
    + "LOS 11 INPUTS EDITABLES (de aqui depende TODA la simulacion):\n"
    + "  A8 Obligacion | B8 Capital | B4 Fecha de desembolso\n"
    + "  C8 Fecha Transferencia | D8 Fecha Vcto Neto | E8 Fecha Vcto Total\n"
    + "  F8 Tasa Negocio(EA) | L8 Tasa Int. Remuneratorio(EA) | AB8 Tasa DPP\n"
    + "  AD8 Parametro GMF | AE8 Parametro Rte Fte\n"
    + "TODO LO DEMAS EN LA FILA 8 ES FORMULA (NO editar): I8 Valor Desembolso, J8 Valor Descuento, O8 Valor Futuro, P8 Intereses totales, Q8 Saldo Capital, S8 Fecha Actual(HOY), Z8 Saldo Int.Rem, AA8 Saldo Int.Mora, AC8 Dcto Pronto Pago, AF8 Valor GMF, AG8 Valor Rte Fte, AH8 Valor Giro.\n\n"
    + "CRONOGRAMA DIARIO (un solo eje de fecha = columna A; A14=B4 y cada fila suma 1 dia):\n"
    + "  Fase Amortizacion (cols A:J): desde el desembolso hasta Vcto Neto.\n"
    + "  Fase Remuneratorio (cols L:Q): desde Vcto Neto hasta Vcto Total. P=causacion rem diaria, Q=int. rem acumulado.\n"
    + "  Fase Moratorio (cols S:W): desde Vcto Total hasta la Fecha Actual. V=causacion mora diaria, W=int. mora acumulado.\n"
    + "  Abonos (cols Y:AB): Y=Abono a capital, Z=Abono a int. rem, AA=Abono a int. mora, AB=Total (FORMULA =SUM(Y:AA)).\n"
    + "  El eje A se extiende al futuro; la fila VIGENTE es donde A = Fecha Actual (S8 = HOY).\n\n"
    + "FLUJO DE TRABAJO:\n"
    + "1) Empieza con get_overview(sheet): te da inputs, resultados, fases, fila vigente y abonos en UNA llamada.\n"
    + "2) Para cambiar un input usa write_named_field (rechaza celdas con formula).\n"
    + "3) Para abonar usa apply_payment: escribe SOLO Y/Z/AA; el Total (AB) se recalcula solo. Si el usuario no discrimina, todo va a capital. Sin fecha, usa la fila vigente (HOY).\n"
    + "4) Para mostrar el cronograma usa get_schedule. Para buscar una fecha, find_date_row.\n"
    + "5) Para cualquier suma/promedio/conteo usa calculate. NUNCA calcules a mano.\n"
    + "6) Para tasa de usura usa read_usura_rate.\n"
    + "7) read_cells/get_cell_info/write_cells solo si lo especifico no alcanza. write_cells NO pisa formulas salvo force=true.\n\n"
    + "FORMATOS: Fechas dd/mm/yyyy. Tasas: 15% = 0.15 (si escribes 16.2345 se interpreta 0.162345). Montos sin separadores.\n\n"
    + "REGLAS:\n"
    + "- NUNCA inventes datos: si no lo leiste con una tool, no lo afirmes.\n"
    + "- NUNCA escribas en celdas con formula ni en hojas protegidas.\n"
    + "- Para abonar nunca toques AB (Total): es formula.\n"
    + "- Tras modificar, reporta hoja, celda(s), fila, valores y el impacto recalculado (saldos, Valor Giro).\n"
    + "- Tablas en markdown:\n"
    + "  | Fecha | Valor |\n"
    + "  |-------|-------|\n"
    + "  | 10/05/2026 | 208.45 |\n\n"
    + "ESTILO:\n"
    + "- Espanol, conciso, directo al resultado.\n"
    + "- Sin saludos ni despedidas (Listo, Perfecto, Claro, Entendido). Sin disculpas.\n"
    + "- No repitas la pregunta ni expliques lo obvio. Ve al dato: tabla, numero o confirmacion.\n"
    + "- Si el usuario dice 'a tu concepto': ACTUA con defaults razonables, sin preguntar.\n"
    + "- Tu ejecutas; no le des instrucciones al usuario.\n"
    + "- Si calculate da 0, es 0. No asumas error.";
}

// ===============================================================
//  UTILIDADES
// ===============================================================

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function safeJsonParse(text) { try { return JSON.parse(text); } catch (e) { return null; } }
function normalizeText(text) { return String(text || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim(); }
function normalizeKey(text) { return normalizeText(text).replace(/\./g, "").replace(/\s+/g, " ").trim(); }
function isBlank(v) { return v === "" || v === null || v === undefined; }

function pickFromRow(rowData, colIdx, cols) {
  for (var i = 0; i < cols.length; i++) {
    var v = rowData[colIdx[cols[i]]];
    if (typeof v === "number" && !isNaN(v)) return formatValue(v);
  }
  return "vacio";
}

function pick(sheet, row, c1, c2, c3) {
  var cols = [c1, c2, c3];
  for (var i = 0; i < cols.length; i++) {
    if (!cols[i]) continue;
    var v = sheet.getRange(row, cols[i]).getValue();
    if (typeof v === "number" && !isNaN(v)) return formatValue(v);
  }
  return "vacio";
}

function toNum(v) {
  if (typeof v === "number") return isNaN(v) ? 0 : v;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

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

// -- Utilidades de fecha --

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

function findRowByDateObj(sheet, dateCol, startRow, lastRow, target) {
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

function findRowByDate(sheet, dateCol, startRow, lastRow, dateStr) {
  return findRowByDateObj(sheet, dateCol, startRow, lastRow, parseDateFlexible(dateStr));
}

function inferHeaderForCell(sheet, schema, row, col) {
  if (row === schema.mainValueRow) {
    for (var k in schema.fields) if (schema.fields[k].col === col && schema.fields[k].row === row) return schema.fields[k].label;
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
    "transferencia": "fecha_transferencia",
    "fecha vcto neto": "fecha_vcto_neto",
    "vcto neto": "fecha_vcto_neto",
    "fecha vencimiento neto": "fecha_vcto_neto",
    "vencimiento neto": "fecha_vcto_neto",
    "fecha vcto total": "fecha_vcto_total",
    "vcto total": "fecha_vcto_total",
    "fecha vencimiento total": "fecha_vcto_total",
    "vencimiento total": "fecha_vcto_total",
    "tasa negocio": "tasa_negocio",
    "tasa spread proveedor": "tasa_negocio",
    "tasa proveedor": "tasa_negocio",
    "spread": "tasa_negocio",
    "tasa remuneratorio": "tasa_remuneratorio",
    "tasa interes remuneratorio": "tasa_remuneratorio",
    "tasa ea interes remuneratorio": "tasa_remuneratorio",
    "remuneratorio": "tasa_remuneratorio",
    "tasa dpp": "tasa_dpp",
    "dpp": "tasa_dpp",
    "gmf": "parametro_gmf",
    "parametro gmf": "parametro_gmf",
    "rte fte": "parametro_rte_fte",
    "retefuente": "parametro_rte_fte",
    "rte": "parametro_rte_fte",
    "parametro rte fte": "parametro_rte_fte"
  };
  return aliases[k] || k;
}
