# Generador de simulaciones por obligación — Angular

Aplicación web Angular para cargar una plantilla Excel, tomar la hoja `Base` como referencia y crear una copia por cada obligación. La hoja `Base` original se conserva intacta y las fórmulas/estilos se copian a cada nueva pestaña.

## Qué hace

- Carga una plantilla `.xlsx` con hoja obligatoria `Base`.
- Puede leer obligaciones desde:
  - la hoja `Obligaciones` dentro de la misma plantilla, o
  - un archivo externo `.xlsx` / `.csv`.
- Crea una hoja por cada obligación.
- Nombra cada hoja con el ID/número de obligación.
- Inserta los datos en las celdas solicitadas:
  - `A8` → Obligación
  - `B8` → Capital
  - `B4` → Fecha de desembolso
  - `C8` → Fecha de Transferencia
  - `D8` → Fecha Vcto Neto
  - `E8` → Fecha Vcto Total
  - `F8` → Tasa negocio
  - `L8` → Tasa interés remuneratorio
  - `AB8` → Tasa DPP
- Convierte tasas tipo `16.6544` a decimal `0.166544` dividiendo entre `100`.
- Fuerza recálculo al abrir el archivo en Excel.

## Requisitos

- Node.js 18+ recomendado.
- npm 9+.

## Instalación

```bash
npm install
npm start
```

Luego abre el navegador en:

```text
http://localhost:4200
```

## Uso

1. Carga tu plantilla `.xlsx` con hoja `Base`.
2. Opcionalmente carga un archivo `.xlsx` o `.csv` con obligaciones.
   - Si no cargas este archivo, la app usará la hoja `Obligaciones` de la plantilla.
3. Mantén activa la opción **Dividir tasas entre 100** si tus tasas vienen como `16.6544`.
4. Haz clic en **Generar Excel**.
5. Se descarga un nuevo `.xlsx` con una pestaña por obligación.

## Encabezados aceptados

La app reconoce encabezados con o sin tildes y algunas variaciones comunes:

- `Obligación`
- `Capital`
- `Fecha de desembolso`
- `Fecha de Transferencia`
- `Fecha Vcto Neto`
- `Fecha Vcto Total`
- `Tasa/Spread proveedor (Tasa Negocio)` o `Tasa negocio`
- `Tasa EA Interés Remuneratorio` o `Tasa interés remuneratorio`
- `Tasa DPP`

También ignora automáticamente una fila auxiliar como la de tu archivo original que contiene referencias `A8`, `B8`, `B4`, etc.

## Buenas prácticas incluidas

- Servicio dedicado para la lógica Excel: `ExcelObligationsService`.
- Modelos tipados en `obligation.model.ts`.
- UI standalone Angular con señales (`signal`).
- Estilos tipo macOS/glassmorphism.
- Sin backend: el procesamiento ocurre localmente en el navegador.
- No modifica la hoja `Base`; genera hojas nuevas a partir de ella.

## Limitaciones importantes

- ExcelJS conserva fórmulas, estilos, merges y formatos principales; elementos muy específicos como macros VBA no se procesan en el navegador.
- Excel calcula las fórmulas al abrir el archivo generado. Por eso se activa `fullCalcOnLoad`.
- Los nombres de hojas de Excel tienen máximo 31 caracteres y no aceptan `\ / * ? : [ ]`; la app sanea esos caracteres automáticamente.
