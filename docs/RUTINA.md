# Rutina diaria de facturas Canavida

Cada día a las 7:05 pm (hora de Costa Rica) un agente:

1. Busca en Gmail (admicanavida@gmail.com) los correos nuevos con PDF reenviados desde factura@canavida.com
   (`has:attachment filename:pdf newer_than:1d`). Si el correo es muy pesado, usa el `.eml` guardado en Google Drive.
2. Lee cada PDF/XML una por una: proveedor, fecha, número, moneda, subtotal, IVA y total.
3. Acomoda los PDF por mes (`AAAA-MM/AAAA-MM-DD_Proveedor_Numero.pdf`).
4. Genera un PDF de resumen y un Excel (hojas `Facturas` y `Resumen por mes`) con `scripts/`.
5. Actualiza el Excel en la carpeta de Drive "Facturas Octubre" (id 1b6a_THjWMSwjlJJSKU2erovRS3hXyNLy).
6. Envía el PDF y el Excel a admicanavida@gmail.com. Si no hay facturas nuevas, envía un aviso corto.

Reglas: recibos, comprobantes de pago y pre-facturas se listan pero no se suman; los contratos del ICE se listan con la
descripción "Contratos del ICE"; los duplicados (misma clave numérica) se cuentan una vez.

Seguridad: NUNCA guardar contraseñas en este repositorio ni en el prompt. El agente usa solo los conectores de Gmail y Drive.
Tampoco se suben facturas (PDF) a este repositorio.

Scripts de referencia (`scripts/`): `extraer_adjuntos.py` -> `leer_facturas.py` -> `generar_reporte.py`.
Usar `/usr/bin/python3` (tiene openpyxl y reportlab).

## Reglas de guardado (actualizado)

- Cada factura nueva se guarda en Drive en la carpeta del MES DE SU FECHA DE EMISION (no la fecha de hoy):
  - Octubre: `Facturas Octubre` (id 1b6a_THjWMSwjlJJSKU2erovRS3hXyNLy)
  - Noviembre: `Facturas Noviembre` (id 1IwSc2GzhS_XNV-4tHPZdnShIjWJM16ud)
  - Otro mes: crear `Facturas <Mes>` en Mi unidad.
  - Nombre del archivo: `AAAA-MM-DD_Proveedor_Numero.pdf`. Verificar el tamaño después de subir.
- Cada factura nueva se agrega como una fila en el libro de Google
  `Canavida Gastos Ingresos Oficial - Octubre (facturas llenadas)` (id 1tjn9HNQOPJhUCz92T-N6jikCVdWXVKTPvPOzFQITS_w),
  hoja `Gastos Diarios`, en la primera fila vacía después de los datos.
  - La fecha de la fila debe ser la fecha de emisión de la factura.
  - Monto en USD; si la factura es en colones se divide entre 459 (tipo de cambio usado por Canavida desde el 10/10/2026).
  - La categoría debe ser una de la hoja `Listas`; si hay duda se marca "Revisar categoría" en la columna H.
  - No duplicar facturas ya registradas (buscar el número en la columna G).
