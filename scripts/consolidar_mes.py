"""Consolida las facturas procesadas por las corridas diarias anteriores (ramas claude/upbeat-bohr-*)
en un solo libro oficial + reporte mensual (Excel y PDF).

Uso: python3 scripts/consolidar_mes.py registros.json AAAA-MM-DD
registros.json: lista con fecha, cat, desc, prov, metodo, crc|usd, ref, nota, pdf, corrida.
"""
import json, sys, datetime as dt, openpyxl
from openpyxl.styles import Font, PatternFill
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont("DV", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

TC = 459
XLSX = "Copy of Canavida Gastos Ingresos Oficial (2).xlsx"
regs = json.load(open(sys.argv[1])); hoy = sys.argv[2]
for x in regs:
    if "crc" in x: x["usd"] = round(x["crc"] / TC, 2)
    else: x["crc"] = round(x["usd"] * TC, 2)
regs.sort(key=lambda v: (v["fecha"], v["prov"]))
nuevos = [x for x in regs if x.get("corrida") == (sys.argv[4] if len(sys.argv) > 4 else None)]

# 1) Libro oficial: filas en Gastos Diarios (sin duplicar por N° factura)
wb = openpyxl.load_workbook(XLSX); ws = wb["Gastos Diarios"]
ya = {str(ws.cell(r, 7).value) for r in range(6, ws.max_row + 1)}
row = 135
for x in regs:
    if x["ref"] in ya: continue
    while ws.cell(row, 2).value is not None: row += 1
    monto = f"={x['crc_formula']}/{TC}" if x.get("crc_formula") else (f"={x['crc']}/{TC}" if x.get("src") == "crc" else x["usd"])
    vals = [dt.datetime.fromisoformat(x["fecha"]), x["cat"], x["desc"], x["prov"], x["metodo"], monto, x["ref"], x.get("nota")]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row, c, v)
        if c == 1: cell.number_format = "dd/mm/yyyy"
        if c == 6: cell.number_format = "#,##0.00"
    x["fila"] = row; row += 1
wb.save(XLSX)

# 2) Excel de reporte
h = ["Mes", "Fecha", "Proveedor", "Categoría", "Descripción", "Método", "N° factura", "Monto ₡", "Monto USD", "Fila Gastos Diarios", "PDF", "Nota"]
w2 = openpyxl.Workbook(); s = w2.active; s.title = "Facturas"; s.append(h)
for c in s[1]: c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDDDDD")
for x in regs:
    s.append([x["fecha"][:7], dt.datetime.fromisoformat(x["fecha"]), x["prov"], x["cat"], x["desc"], x["metodo"], x["ref"],
              x["crc"], x["usd"], x.get("fila"), x.get("pdf"), x.get("nota")])
for r in s.iter_rows(min_row=2):
    r[1].number_format = "dd/mm/yyyy"; r[7].number_format = r[8].number_format = "#,##0.00"
for col, wd in zip("ABCDEFGHIJKL", [9, 11, 30, 26, 50, 20, 18, 14, 12, 10, 60, 40]): s.column_dimensions[col].width = wd
s2 = w2.create_sheet("Resumen por mes"); s2.append(["Mes", "Categoría", "Documentos", "Monto ₡", "Monto USD"])
for c in s2[1]: c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDDDDD")
agg, mes = {}, {}
for x in regs:
    a = agg.setdefault((x["fecha"][:7], x["cat"]), [0, 0, 0]); a[0] += 1; a[1] += x["crc"]; a[2] += x["usd"]
    m = mes.setdefault(x["fecha"][:7], [0, 0, 0]); m[0] += 1; m[1] += x["crc"]; m[2] += x["usd"]
for (m, cat), (n, crc, usd) in sorted(agg.items()): s2.append([m, cat, n, round(crc, 2), round(usd, 2)])
for m, (n, crc, usd) in sorted(mes.items()):
    s2.append([m, "TOTAL " + m, n, round(crc, 2), round(usd, 2)])
    for c in s2[s2.max_row]: c.font = Font(bold=True)
for r in s2.iter_rows(min_row=2): r[3].number_format = r[4].number_format = "#,##0.00"
for col, wd in zip("ABCDE", [9, 32, 12, 15, 12]): s2.column_dimensions[col].width = wd
xl = f"reportes/facturas_{hoy}.xlsx"; w2.save(xl)

# 3) PDF de reporte
st = getSampleStyleSheet()
for k in ("Normal", "Title", "Heading3"): st[k].fontName = "DV" if k == "Normal" else "DVB"
sm = st["Normal"].clone("sm"); sm.fontSize = 7; sm.leading = 8.5
fm = lambda v: f"{v:,.2f}"
fd = dt.date.fromisoformat(hoy).strftime("%d/%m/%Y")
ts = lambda fs: TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("GRID", (0, 0), (-1, -1), .3, colors.grey),
    ("FONTNAME", (0, 0), (-1, -1), "DV"), ("FONTSIZE", (0, 0), (-1, -1), fs), ("VALIGN", (0, 0), (-1, -1), "TOP")])
d2 = [["Mes", "Categoría", "Docs", "Monto ₡", "USD"]] + [[m, c, n, fm(a), fm(u)] for (m, c), (n, a, u) in sorted(agg.items())]
d2 += [[m, "TOTAL " + m, n, fm(a), fm(u)] for m, (n, a, u) in sorted(mes.items())]
t2 = Table(d2, colWidths=[50, 170, 40, 90, 70]); t2.setStyle(ts(8))
t2.setStyle(TableStyle([("ALIGN", (2, 1), (-1, -1), "RIGHT"), ("FONTNAME", (0, len(d2) - len(mes)), (-1, -1), "DVB")]))
d1 = [["Fecha", "Proveedor", "Categoría", "Descripción", "Método", "N°", "₡", "USD"]] + [
    [x["fecha"], Paragraph(x["prov"], sm), Paragraph(x["cat"], sm), Paragraph(x["desc"] + (f" <i>({x['nota']})</i>" if x.get("nota") else ""), sm),
     Paragraph(x["metodo"], sm), Paragraph(x["ref"], sm), fm(x["crc"]), fm(x["usd"])] for x in regs]
t1 = Table(d1, repeatRows=1, colWidths=[55, 105, 90, 250, 70, 75, 70, 55]); t1.setStyle(ts(7))
t1.setStyle(TableStyle([("ALIGN", (6, 1), (7, -1), "RIGHT")]))
nov = mes.get(hoy[:7], [0, 0, 0])
pdf = f"reportes/resumen_facturas_{hoy}.pdf"
SimpleDocTemplate(pdf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24).build([
    Paragraph(f"Canavida · Actualización de facturas {fd}", st["Title"]),
    Paragraph(f"<b>Hoy: {len(nuevos)} facturas PDF nuevas</b>" + (f" ({', '.join(x['prov'] + ' ' + x['ref'] for x in nuevos)}: ₡{fm(sum(x['crc'] for x in nuevos))} ≈ US${fm(sum(x['usd'] for x in nuevos))})" if nuevos else "") + f". Acumulado {hoy[:7]} en el repositorio: "
              f"{nov[0]} documentos · ₡{fm(nov[1])} ≈ US${fm(nov[2])} (TC ₡{TC}). PDF guardados en facturas/AAAA-MM por mes de emisión; "
              f"filas registradas en la hoja 'Gastos Diarios' del libro oficial.", st["Normal"]),
    Spacer(1, 4), Paragraph(sys.argv[3] if len(sys.argv) > 3 else "", st["Normal"]),
    Spacer(1, 8), Paragraph("Resumen por mes y categoría", st["Heading3"]), t2, Spacer(1, 10),
    Paragraph("Detalle de documentos", st["Heading3"]), t1])
json.dump(regs, open(f"reportes/facturas_{hoy}.json", "w"), ensure_ascii=False, indent=1, default=str)
print(xl, pdf, {k: [v[0], round(v[1], 2), round(v[2], 2)] for k, v in mes.items()})
