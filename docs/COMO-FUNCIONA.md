# Cómo funciona el sistema

## Arquitectura general

```
┌──────────────────────────┐     ┌─────────────────────┐
│   Angular 17 (Frontend)  │────▶│  Google Apps Script  │
│   Host: Vercel (gratis)  │     │  (middleware/server) │
└──────────┬───────────────┘     └─────────┬───────────┘
           │                               │
           │ ExcelJS (lectura/escritura)    │ Google Sheets API
           │ PapaParse (CSV)               │ DeepSeek API (chat IA)
           │ FileSaver (descarga)          │ Google Drive (upload)
           ▼                               ▼
    ┌──────────────┐              ┌─────────────────┐
    │  Navegador   │              │  Google Sheets   │
    │  (procesa en │              │  (cálculos,      │
    │   local)     │              │   fórmulas,      │
    └──────────────┘              │   preview)       │
                                  └─────────────────┘
                                         │
                                  ┌──────┴──────────┐
                                  │  Supabase        │
                                  │  (historial,     │
                                  │   presets,       │
                                  │   preferencias)  │
                                  └─────────────────┘
```

## Flujo completo

### 1. Carga de plantilla (Excel base)
- Usuario arrastra un `.xlsx` con una hoja obligatoria llamada `Base`
- ExcelJS lo lee en el navegador (todo local)
- La hoja `Base` contiene: celdas de entrada (filas 4-8), fórmulas financieras (filas 14+), formato y estilos

### 2. Carga de obligaciones
- **Opción A**: hoja `Obligaciones` dentro de la misma plantilla → ExcelJS lee las filas
- **Opción B**: archivo externo `.xlsx` o `.csv` → ExcelJS o PapaParse
- **Opción C**: pegar CSV desde portapapeles
- Campos mapeados: Obligación, Capital, Fechas, Tasas
- Compatible con tildes y variaciones de encabezados

### 3. Generación del Excel
- Por cada obligación se clona la hoja `Base` (fórmulas, estilos, merge cells)
- Se insertan los datos en celdas fijas: A8, B8, B4, C8, D8, E8, F8, L8, AB8
- Tasas se dividen entre 100 si está activado (16.65 → 0.1665)
- Fechas se convierten a serial de Excel (sin el bug del +1 día)
- `fullCalcOnLoad = true` para que Excel recalcule al abrir
- El archivo se descarga con FileSaver

### 4. Subida a Google Drive (opcional)
- El `.xlsx` generado se codifica en Base64
- Se envía a Google Apps Script vía POST
- Apps Script lo convierte a Google Sheets nativo
- Retorna URL de vista previa (iframe embebido en la app)

### 5. Chat con IA (asistente DeepSeek)
- El usuario escribe preguntas en lenguaje natural
- El frontend envía `{ action: "chat", id: sheetId, history, question }` a Apps Script
- Apps Script construye un prompt de sistema con la estructura de las hojas
- Llama a DeepSeek API (`deepseek-chat`) con function calling
- El modelo decide qué herramientas usar en cada turno:

| Tool | Qué hace |
|------|----------|
| `list_sheets` | Lista hojas de simulación (omite Base, Obligaciones) |
| `read_cells` | Lee rangos del sheet real |
| `get_cell_info` | Lee valor + fórmula + nota de una celda |
| `write_cells` | Escribe/limpia valores (abonos, tasas, fechas) |
| `calculate` | Ejecuta `=SUM()`, `=AVERAGE()` en hoja oculta `_scratch` |

- Hasta 20 turnos de tool-calling por consulta
- La respuesta se devuelve como texto markdown
- El frontend convierte markdown → HTML (`###`, tablas, **bold**, etc.)

### 6. Historial y presets (Supabase)
- Cada simulación subida se registra: id, nombre, fecha, URL del sheet
- El usuario puede ver historial, reabrir sheets anteriores, eliminar
- Presets de prompts: frases predefinidas que se cargan en el chat (sin auto-enviar)

## Datos que no salen del navegador
- Lectura de archivos Excel/CSV
- Generación del Excel (ExcelJS)
- Conversión markdown → HTML del chat

## Datos que pasan por internet
- Subida del Excel a Google Drive (Base64)
- Consultas al chat IA (pregunta + historial)
- CRUD de historial/presets en Supabase
