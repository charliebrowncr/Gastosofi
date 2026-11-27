# Rutina diaria: facturas de Gmail → plantilla Canavida

Cada corrida programada sigue estos pasos.

1. **Buscar** en Gmail (cuenta admicanavida@gmail.com, conector Gmail; no se usan contraseñas) con
   `filename:pdf after:<fecha de ayer> -in:sent`. Si un día no se envió reporte, usar como `after:` la fecha del último
   reporte de facturas enviado (`in:sent subject:facturas`) para no perder las facturas de los días sin corrida.
2. **Descartar** los IDs de mensaje que ya están en la columna `message_id` de `data/facturas.csv`.
3. **Descargar** de cada mensaje nuevo sus adjuntos `.pdf` y los `.xml` de la factura (no los de respuesta de Hacienda) en
   `_inbox/<message_id>/`, junto con un `meta.json` (`id`, `date`, `subject`, `body`).
4. **Procesar** con `python3 scripts/procesar_facturas.py --base data/base.xlsx AAAA-MM-DD`, donde `data/base.xlsx` es el Excel
   adjunto en el reporte del día anterior (descargarlo de Gmail). Así se conservan las filas "Auto Gmail" previas aunque el
   contenedor no guarde `data/`. Si la factura viene solo en PDF, se puede poner `"proveedor"` en `meta.json`. El script:
   - lee el XML de Hacienda, o el texto del PDF si no hay XML;
   - asigna la categoría según los registros que ya están en la hoja y unas reglas por palabra clave;
   - convierte los montos en colones a dólares con 459 (`TIPO_CAMBIO`);
   - deja fuera los duplicados, lo que ya estaba en la hoja, las pre-facturas y lo emitido por Canavida;
   - archiva cada PDF en `facturas/AAAA-MM/`;
   - genera `reportes/Canavida_Gastos_Ingresos_<fecha>.xlsx` (la plantilla con las filas "Auto Gmail" en "Gastos Diarios")
     y `reportes/Reporte_Facturas_<fecha>.pdf`.
5. **Enviar** a admicanavida@gmail.com un correo con los dos archivos y un resumen.
6. Los datos (`data/`, `facturas/`, `reportes/`) **no se suben a GitHub** porque contienen información financiera. Como el contenedor es efímero, la corrida siguiente evita duplicados filtrando por fecha (`after:` = día anterior) y con la clave de Hacienda.

## Notas
- La plantilla original no se modifica. El Excel se vuelve a generar cada día a partir de la plantilla y de `data/facturas.csv`.
- En "Gastos e Ingresos", la columna de octubre tiene valores fijos. Noviembre y diciembre suman solos desde "Gastos Diarios".
- Las filas con "Revisar categoría" o con moneda asumida conviene confirmarlas a mano.
