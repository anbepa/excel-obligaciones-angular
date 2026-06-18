# Google Apps Script — Documentación completa

## Configuración global

```
Línea 1:    DEEPSEEK_API_KEY  → desde PropertiesService (no hardcodeada)
Línea 2:    DEFAULT_FOLDER_ID → carpeta Drive donde se suben los Excel
Línea 3:    MAX_TOOL_TURNS    → 20 (máx ida/vuelta DeepSeek por consulta)
```

---

## Las 5 TOOLS (Function Calling)

DeepSeek recibe estas definiciones y decide autónomamente cuál usar. Cada tool tiene nombre, descripción y parámetros tipados.

```
┌───────────────┬──────────────────────────────────────────────────────┐
│ Tool          │ Qué hace                                             │
├───────────────┼──────────────────────────────────────────────────────┤
│ list_sheets   │ Lista hojas (omite Base, Obligaciones, Tasa usura).  │
│               │ Retorna JSON [{name, rows, cols}]                    │
├───────────────┼──────────────────────────────────────────────────────┤
│ read_cells    │ Lee rango A1 del sheet real. Máx 200 filas.          │
│               │ Retorna formato compacto: "r14:A=03/04 | B=1995 |..."│
│               │ Si es 1 celda: "1995" o "1995 (fórmula: =B8*2)"     │
├───────────────┼──────────────────────────────────────────────────────┤
│ get_cell_info │ Lee valor + fórmula + nota de UNA celda.             │
│               │ Retorna: "Celda: B8\nFórmula: =SUM(A:A)\nValor: 3M" │
├───────────────┼──────────────────────────────────────────────────────┤
│ write_cells   │ Escribe valores en rango. Normaliza fechas y comas.  │
│               │ Si todos vacíos → clearContent().                    │
│               │ Retorna: "✓ Escrito en Hoja!Y148:AB148 → [[...]]"   │
├───────────────┼──────────────────────────────────────────────────────┤
│ calculate     │ Escribe fórmula en hoja oculta _scratch,            │
│               │ lee resultado con getValue(), limpia.                │
│               │ Retorna: "Resultado de =SUM(V:V) en Hoja = 152.715" │
└───────────────┴──────────────────────────────────────────────────────┘
```

### Ejemplo: toolReadCells

La respuesta que le llega a DeepSeek tiene este formato:

```
r106:A=03/04/2026 | B=3000000 | V=vacio
r107:A=04/04/2026 | B=2998005 | V=1995
r108:A=05/04/2026 | B=2996010 | V=1995
...
```

- Solo columnas con al menos 1 valor no vacío
- `vacio` = celda sin valor ni fórmula
- Si hay fórmula se incluye: `B=3000000 (fórmula: =B8-Y106)`
- Truncado a 200 filas máximo

### Ejemplo: toolCalculate

```
1. Obtiene/crea hoja oculta "_scratch"
2. scratch.getRange("A1").setFormula("=SUM(V106:V181)")
3. SpreadsheetApp.flush()       ← fuerza recálculo
4. result = scratch.getRange("A1").getValue()
5. scratch.getRange("A1").clearContent()
6. Retorna: "Resultado de =SUM(V106:V181) en AAC324829 = 152.715"
```

---

## Endpoints HTTP

### `doPost(e)` — Línea 89

```
POST /exec
Body: { action?: "chat", id?, question?, history?, fileBase64?, fileName?, folderId? }

Si action === "chat" → handleChat(data)
Si no                       → uploadExcel(data)
```

**Flujo de subida de Excel:**

```
1. Valida fileBase64 + fileName
2. Verifica carpeta Drive (DEFAULT_FOLDER_ID)
3. Decodifica Base64 → Blob .xlsx
4. Crea archivo Excel en Drive
5. Convierte a Google Sheets nativo:
   Drive.Files.create({ mimeType: "google-apps-spreadsheet" }, blob)
6. Borra el .xlsx original (setTrashed)
7. Comparte con ANYONE_WITH_LINK + EDIT
8. Retorna { success, id: sheetId, url: "https://docs.google.com/..." }
```

### `doGet(e)` — Línea 137

```
GET /exec?action=delete&id=SHEET_ID
→ DriveApp.getFileById(id).setTrashed(true)
→ { success: true }
```

---

## Chat con Function Calling — `handleChat()` Línea 153

### Constructor de mensajes (línea 170-177)

```
messages = [
  { role: "system", content: buildSystemPrompt() },   ← mapa de la hoja, reglas
  ...historial previo (user + assistant, sin system),  ← contexto de conversación
  { role: "user", content: pregunta }                  ← nueva pregunta
]
```

### Loop de tool calling (línea 184-248)

```
while (turn < 20) {
    turn++

    ┌─ POST a https://api.deepseek.com/v1/chat/completions
    │  payload: { model: "deepseek-chat", messages, tools: TOOLS, temperature: 0.1 }
    │
    ▼
    ¿result.choices[0].message.tool_calls?

    SÍ ──→ Para cada tool_call:
    │      1. Guardar msg del asistente en historial
    │      2. Ejecutar executeTool(ss, fnName, fnArgs)
    │      3. Guardar resultado como { role: "tool", content: resultado }
    │      4. continue (siguiente turno)
    │
    NO ──→ finish_reason === "stop"
           finalReply = msg.content
           break
}

Retorna { success: true, reply: finalReply }
```

### Manejo de errores en el loop

```
- Si result.error → "DeepSeek: mensaje de error"
- Si 20 turnos sin respuesta → "El asistente no generó respuesta después de 20 turnos"
- Si catch general → error.toString()
```

---

## Ejecución de Tools — `executeTool()` Línea 273

Router simple que redirige a la función correspondiente:

```
list_sheets   → toolListSheets(ss)
get_cell_info → toolGetCellInfo(ss, sheet, cell)
read_cells    → toolReadCells(ss, sheet, range)
write_cells   → toolWriteCells(ss, sheet, range, values)
calculate     → toolCalculate(ss, sheet, formula)
otro          → "ERROR: herramienta desconocida"
```

### `toolListSheets(ss)` — Línea 296

```
1. ss.getSheets() → itera todas las hojas
2. Normaliza nombre (sin tildes, lowercase)
3. Omite: "base", "obligaciones", hojas con "tasa usura"
4. Retorna JSON: [{"name":"AAC324829","rows":181,"cols":28}]
```

### `toolGetCellInfo(ss, sheet, cell)` — Línea 315

```
1. findSheet() → busca por nombre exacto o aproximado
2. parseRange("B8") → { startRow:8, startCol:2, numRows:1, numCols:1 }
3. range.getValue() + range.getFormula() + range.getNote()
4. Retorna multilínea:
   Celda: B8
   Fórmula: =SUM(A14:A200)
   Valor calculado: 3000000
   Nota: Capital inicial
```

### `toolWriteCells(ss, sheet, range, values)` — Línea 404

```
1. findSheet() → valida hoja
2. parseRange() → valida rango
3. Normaliza cada valor con normalizeInputValue():
   - "15/05/2026" → new Date(2026,4,15)   ← string fecha
   - "1.500,50"   → 1500.5                ← número con coma decimal
   - 500000       → 500000                ← número puro
4. Si todos los valores son "" → range.clearContent()
5. Si no → range.setValues(writeValues)
6. SpreadsheetApp.flush() → fuerza recálculo de fórmulas
```

### `toolCalculate(ss, sheet, formula)` — Línea 459

```
1. Valida que la fórmula empiece con "="
2. ss.getSheetByName("_scratch") o ss.insertSheet("_scratch").hideSheet()
3. scratch.getRange("A1").setFormula("=SUM(V106:V181)")
4. SpreadsheetApp.flush()
5. result = scratch.getRange("A1").getValue()
6. scratch.getRange("A1").clearContent()
7. Retorna resultado formateado
```

---

## Utilidades compartidas

### `parseRange("A14:J50")` — Línea 493

```
Regex: /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i

"A14:J50" → { startCol:1, startRow:14, endCol:10, endRow:50, numRows:37, numCols:10 }
"B8"     → { startCol:2, startRow:8,  endCol:2,  endRow:8,  numRows:1,  numCols:1  }
"AB14"   → { startCol:28, startRow:14, ... }
```

### `colLetterToNumber("AB")` — Línea 523

```
A=1, B=2, ..., Z=26, AA=27, AB=28, ..., ZZ=702
Algoritmo: col = col * 26 + (charCode - 64)
```

### `columnToLetter(27)` — Línea 662

```
0=A, 1=B, ..., 25=Z, 26=AA, 27=AB, ...
Inverso de colLetterToNumber
```

### `findSheet(ss, name)` — Línea 637

```
1. Búsqueda exacta por nombre
2. Búsqueda normalizada (sin tildes, lowercase)
3. Búsqueda por subcadena
```

### `normalizeInputValue(value)` — Línea 623

```
"15/05/2026"          → new Date(2026, 4, 15)    ← detecta fecha dd/mm/yyyy
"1,500,000"           → 1500000                   ← remueve comas de miles
"1.500,50"            → 1500.5                    ← coma decimal a punto
500000                → 500000                    ← número sin cambios
```

### `formatValue(val)` — Línea 652

```
"" / null / undefined → "vacio"
Date                  → "15/05/2026" (formato dd/MM/yyyy)
número entero         → "1995"
número decimal        → "1.995,00" (locale es-CO, 2-6 decimales)
string                → valor original
```

### `normalizeText(text)` — Línea 619

```
"Obligación" → "obligacion"  (NFD, remove diacritics, lowercase, trim)
```

---

## System Prompt — `buildSystemPrompt()` Línea 536

El prompt que DeepSeek recibe como `{ role: "system" }` tiene 7 secciones:

| Sección | Contenido |
|---------|-----------|
| Rol | "Eres un asistente experto en análisis y modificación..." |
| Herramientas | Descripción de las 5 tools y cuándo usar cada una |
| Estructura de hoja | Mapa de filas 1-14: qué hay en cada celda (A8, B8, F8, etc.) |
| Tablas diarias | Las 4 tablas: Amortización, Int Rem, Int Mora, Abonos |
| Modificaciones | Cómo hacer abonos, cambiar tasas, formato de fechas/montos |
| Reglas | No inventar, usar calculate, formato tabla markdown |
| Comportamiento proactivo | No preguntar, actuar con defaults, no dar instrucciones al usuario |
| Estilo de respuesta | Prohibido saludos, disculpas, relleno. Directo al dato. |

---

## Resumen del flujo completo

```
Angular (frontend)
    │
    │ POST /exec
    │ { action: "chat", id, history, question }
    ▼
Google Apps Script
    │
    ├─ doPost() → handleChat()
    │
    ├─ Construye messages = [system prompt, ...historial, pregunta]
    │
    ├─ Loop (máx 20 turnos):
    │   │
    │   ├─ POST DeepSeek API { model, messages, tools, temperature }
    │   │
    │   ├─ DeepSeek decide: ¿usar tool o responder?
    │   │
    │   ├─ Si tool_calls:
    │   │   executeTool(ss, toolName, args)
    │   │   → getRange().getValues() / getFormula() / setValues() / setFormula()
    │   │   → resultado se agrega al historial
    │   │   → siguiente turno
    │   │
    │   └─ Si no: respuesta final → break
    │
    └─ Retorna { success: true, reply: "texto markdown" }
        │
        ▼
    Angular: mdToHtml(reply) → renderiza en chat
```
