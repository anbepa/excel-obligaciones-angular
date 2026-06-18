# Cómo funciona Google Apps Script y la comunicación con el agente IA

## Rol de Google Apps Script

Es el **middleware** entre Angular y Google Sheets + DeepSeek. Actúa como un mini-servidor sin costo que:

1. Recibe peticiones HTTP desde el frontend Angular
2. Convierte Excel a Google Sheets y los sube a Drive
3. Intermedia entre DeepSeek y Google Sheets usando **function calling**

## Endpoints expuestos

| Método | Endpoint | Acción |
|--------|----------|--------|
| `POST` | `/exec` con `action: "chat"` | Chat con IA (function calling) |
| `POST` | `/exec` sin action | Subir Excel y convertir a Google Sheets |
| `GET` | `/exec?action=delete&id=...` | Eliminar archivo de Drive |

## Flujo del chat: Function Calling paso a paso

### Diagrama de la conversación

```
Usuario escribe: "hacer abono de $500,000 el 15/05/2026"
        │
        ▼
┌──────────────────────────────────────────────┐
│ Angular envía POST a Apps Script:            │
│ {                                            │
│   action: "chat",                            │
│   id: "1a2b3c...",     // Google Sheet ID    │
│   history: [...],       // mensajes previos   │
│   question: "hacer abono de $500,000..."     │
│ }                                            │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: handleChat()                    │
│                                              │
│ 1. Construye mensajes:                       │
│    [                                         │
│      { role: "system", content: PROMPT },    │
│      ...historial anterior,                  │
│      { role: "user", content: pregunta }     │
│    ]                                         │
│                                              │
│ 2. Adjunta definición de TOOLS (5 funciones) │
│    - list_sheets                             │
│    - read_cells                              │
│    - get_cell_info                           │
│    - write_cells                             │
│    - calculate                               │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ POST a https://api.deepseek.com/v1/chat/...  │
│                                              │
│ DeepSeek analiza el prompt y decide:          │
│ "Necesito listar las hojas primero y luego    │
│  leer la tabla de abonos para encontrar la   │
│  fila correcta por fecha"                    │
│                                              │
│ Respuesta de DeepSeek (turno 1):             │
│ {                                            │
│   tool_calls: [                              │
│     { function: { name: "list_sheets" } }    │
│   ]                                          │
│ }                                            │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: executeTool(ss, "list_sheets")   │
│                                              │
│ → Lee todas las hojas (omite Base,           │
│   Obligaciones, Tasa usura)                  │
│ → Retorna JSON:                              │
│   [{"name":"AAC324829","rows":181,"cols":28}]│
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ POST a DeepSeek (turno 2)                    │
│                                              │
│ DeepSeek ya sabe la hoja. Ahora pide:        │
│ {                                            │
│   tool_calls: [                              │
│     { function: {                            │
│         name: "read_cells",                  │
│         arguments: {                         │
│           sheet: "AAC324829",                │
│           range: "Y106:Y181"                 │
│         }                                    │
│     }}                                       │
│   ]                                          │
│ }                                            │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: toolReadCells()                  │
│                                              │
│ → Lee columna Y (Abono a Capital) fila 106-181│
│ → Retorna: "r106:vacio | r107:vacio | ..."   │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ POST a DeepSeek (turno 3)                    │
│                                              │
│ DeepSeek ve todo vacío → primera fila libre  │
│ Busca fila por fecha → pide columna A:       │
│ {                                            │
│   tool_calls: [                              │
│     { function: {                            │
│         name: "read_cells",                  │
│         arguments: {                         │
│           sheet: "AAC324829",                │
│           range: "A106:A181"                 │
│         }                                    │
│     }}                                       │
│   ]                                          │
│ }                                            │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: toolReadCells()                  │
│                                              │
│ → Retorna: "r106:03/04/2026 | r107:04/04..." │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ POST a DeepSeek (turno 4)                    │
│                                              │
│ DeepSeek encuentra fila 148 = 15/05/2026     │
│ Decide escribir el abono:                    │
│ {                                            │
│   tool_calls: [                              │
│     { function: {                            │
│         name: "write_cells",                 │
│         arguments: {                         │
│           sheet: "AAC324829",                │
│           range: "Y148:AB148",               │
│           values: [[500000, 0, 0, 500000]]   │
│         }                                    │
│     }}                                       │
│   ]                                          │
│ }                                            │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: toolWriteCells()                 │
│                                              │
│ → sheet.getRange(148, 25, 1, 4).setValues()  │
│ → Google Sheets recalcula TODAS las fórmulas │
│ → Retorna: "✓ Escrito en AAC324829!Y148:AB148"│
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ POST a DeepSeek (turno 5)                    │
│                                              │
│ DeepSeek ya no necesita más tools.           │
│ Genera respuesta final con tabla resumen.    │
│                                              │
│ finish_reason: "stop"                        │
│ content: "Abono aplicado:\n                 │
│           | Campo | Valor |\n..."           │
└──────────────────────────────────────────────┘
        │
        ▼
┌──────────────────────────────────────────────┐
│ Apps Script: retorna la respuesta a Angular  │
│                                              │
│ { success: true, reply: "Abono aplicado..." }│
└──────────────────────────────────────────────┘
        │
        ▼
     Angular muestra la respuesta
```

## Las 5 herramientas (TOOLS)

Cada tool es una función que DeepSeek puede invocar. La definición incluye nombre, descripción y parámetros. DeepSeek **decide autónomamente** cuál usar y con qué argumentos.

### `list_sheets`
```
Entrada:  nada
Salida:   JSON con [{name, rows, cols}, ...]
Uso:      Saber qué hojas de simulación existen
```

### `read_cells`
```
Entrada:  sheet, range (ej: "V106:V181")
Salida:   "r106:1995 | r107:1995 | r108:vacio | ..."
Uso:      Leer datos reales del sheet. Máx 200 filas.
```

### `get_cell_info`
```
Entrada:  sheet, cell (ej: "B8")
Salida:   "Celda: B8\nFórmula: =SUM(...)\nValor calculado: 3000000"
Uso:      Inspeccionar una celda con fórmula (B8, F8, Q8...)
```

### `write_cells`
```
Entrada:  sheet, range, values (ej: [[500000,0,0,500000]])
Salida:   "✓ Escrito en AAC324829!Y148:AB148 → [[500000,0,0,500000]]"
Uso:      Modificar abonos, tasas, fechas. Limpiar con valores vacíos.
```

### `calculate`
```
Entrada:  sheet, formula (ej: "=SUM(V106:V181)")
Salida:   "Resultado de =SUM(V106:V181) en AAC324829 = 152.715"
Uso:      Ejecuta la fórmula en hoja oculta _scratch, lee resultado, limpia.
          Google Sheets hace el cálculo → resultado exacto garantizado.
```

## El loop de tool calling

```javascript
var turn = 0;
while (turn < MAX_TOOL_TURNS) {  // 20 turnos máximo
    turn++;
    
    // 1. Enviar mensajes + tools a DeepSeek
    var payload = {
        model: "deepseek-chat",
        messages: messages,
        tools: TOOLS,
        temperature: 0.1
    };
    
    var response = UrlFetchApp.fetch(DEEPSEEK_URL, options);
    var msg = response.choices[0].message;
    
    // 2. ¿DeepSeek quiere usar una tool?
    if (msg.tool_calls && msg.tool_calls.length > 0) {
        // Agregar respuesta del asistente al historial
        messages.push({ role: "assistant", tool_calls: msg.tool_calls });
        
        // Ejecutar cada tool pedida
        for (var t = 0; t < msg.tool_calls.length; t++) {
            var fnName = msg.tool_calls[t].function.name;
            var fnArgs = msg.tool_calls[t].function.arguments;
            
            var result = executeTool(ss, fnName, fnArgs);
            
            // Agregar resultado al historial
            messages.push({ role: "tool", tool_call_id: ..., content: result });
        }
        continue; // Siguiente turno
    }
    
    // 3. No hay tool_calls → respuesta final
    finalReply = msg.content;
    break;
}
```

Cada turno es una ida y vuelta a la API de DeepSeek. El historial crece con cada tool call y su resultado. DeepSeek usa ese contexto acumulado para decidir el siguiente paso.

## Lo que DeepSeek NO puede hacer

- **No ve el Excel directamente** — solo ve lo que las tools le devuelven
- **No ejecuta fórmulas** — solo ve resultados numéricos (por eso existe `calculate`)
- **No sabe la estructura de las hojas sin leerlas** — por eso el system prompt le da un mapa
- **No modifica sin `write_cells`** — todo cambio pasa por `sheet.getRange().setValues()`
