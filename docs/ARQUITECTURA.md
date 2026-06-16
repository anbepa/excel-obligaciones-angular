# Arquitectura

```text
src/app
├── app.component.*                 # UI principal
├── models
│   └── obligation.model.ts          # Tipos, opciones y mapeo de celdas
└── services
    └── excel-obligations.service.ts # Lectura, copiado y generación Excel
```

## Flujo

1. `AppComponent` recibe archivos desde la interfaz.
2. `ExcelObligationsService` lee la plantilla con ExcelJS.
3. Si hay archivo externo de obligaciones, lo procesa con ExcelJS o PapaParse.
4. Si no hay archivo externo, lee la hoja `Obligaciones` de la plantilla.
5. Por cada registro:
   - crea una hoja nueva,
   - copia `Base`,
   - inserta valores en las celdas configuradas,
   - convierte tasas a decimal si la opción está activa.
6. Descarga el archivo generado con FileSaver.
```
