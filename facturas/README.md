# Facturas desde Gmail → Excel por mes + reporte PDF (corrida diaria)

Cada día, el agente programado:

1. **Recupera el estado anterior**: busca en Gmail el último correo enviado con asunto
   `Reporte diario de facturas Canavida` (`in:sent`) y descarga su adjunto `.xlsx`.
   El registro acumulado de facturas vive en la hoja oculta `_registro` de ese Excel.
   (Si no existe, empieza desde cero.)
2. **Busca facturas nuevas**: `filename:pdf -in:sent after:<fecha del último reporte - 1 día>`.
   Para cada mensaje descarga cada `.pdf` y el `.xml` del comprobante electrónico de Hacienda
   (se omiten los XML de respuesta/aceptación de Hacienda: `AHC_`, `MH_`, `Respuesta`, `MensajeHacienda`)
   en `inbox/<messageId>/` junto con un `meta.json` (`messageId`, `date`, `subject`, `sender`, `bodyExcerpt`, `files`).
3. **Procesa**:

   ```bash
   python3 facturas/procesar_facturas.py \
     --plantilla "Copy of Canavida Gastos Ingresos Oficial (2).xlsx" \
     --inbox <dir>/inbox --salida <dir>/out \
     --previo <Excel del reporte anterior> --tc 459 --fecha AAAA-MM-DD
   ```

   Requiere `openpyxl`, `reportlab` y `pdftotext`/`pdfunite`/`pdfinfo` (poppler).
4. **Envía** a la cuenta de administración un correo con:
   - `Canavida Reporte Facturas AAAA-MM-DD.pdf`: resumen por mes y categoría, facturas nuevas,
     detalle por mes y, anexadas al final, las facturas PDF nuevas ordenadas por fecha
     (solo si todo cabe en 9 MB; el conector de Gmail no admite adjuntos de más de 10 MB).
   - `Canavida Facturas AAAA-MM-DD.xlsx`: la plantilla oficial con
     - una hoja **"Facturas <Mes> <Año>"** por mes (mismas columnas que *Gastos Diarios* + moneda, monto original,
       tipo de cambio, tipo, estado y nombre del PDF);
     - las facturas de gasto nuevas (no duplicadas, del 01/10 al 31/12/2026) agregadas al final de *Gastos Diarios*
       con la nota `Auto: Gmail`, para que el resumen mensual las sume;
     - la hoja oculta `_registro` (estado para el día siguiente).
   Si no hay facturas nuevas, igual se envía el resumen actualizado.

## Reglas

- Los datos (fecha, proveedor, número, total, moneda, medio de pago, detalle) salen del XML de Hacienda;
  si el correo solo trae PDF se leen del texto del PDF y quedan marcados "Datos leídos del PDF (sin XML)".
- Montos en colones → USD con el tipo de cambio `--tc` (fórmula `=monto/tc` igual que en la plantilla).
- Categoría: la misma que se usó para ese proveedor en *Gastos Diarios*; si no existe, por palabras clave;
  por defecto "Imprevistos y otros".
- Duplicados: misma clave de Hacienda o mismo PDF (hash) se ignoran. Se marcan como **Posible duplicado** y no se
  suman: los números de factura ya escritos a mano en *Gastos Diarios*, y los PDF sin XML con el mismo monto que
  otra factura de ±20 días (recibos de pago, reenvíos).
- Se separan: notas de crédito, pre-facturas/cotizaciones, facturas emitidas por Canavida y documentos que no son
  factura (contratos, reservas, hojas de check-in).
- Los PDF y datos financieros **no** se suben a este repositorio (es público); solo vive aquí el código.
