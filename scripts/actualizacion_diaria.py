"""Rutina diaria de facturas Canavida.

Uso: python3 scripts/actualizacion_diaria.py facturas_dia.json AAAA-MM-DD

facturas_dia.json: lista de facturas ya leídas (del XML/PDF de cada correo):
  fecha, tipo, num, prov, desc, cat, metodo, crc (monto en colones), pdf (ruta local),
  nota (opcional), ajuste_crc (opcional: nota de crédito que se resta).
Hace: copia cada PDF a facturas/AAAA-MM/ (mes de emisión), agrega filas a
"Gastos Diarios" del Excel oficial (monto = colones / TC), y genera en reportes/
un Excel y un PDF con el detalle del día y el resumen por mes.
"""
import json, os, re, shutil, sys, datetime as dt
import openpyxl
from openpyxl.styles import Font, PatternFill
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont("DV", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DVB", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
pdfmetrics.registerFontFamily("DV", normal="DV", bold="DVB", italic="DV", boldItalic="DVB")

TC = 459  # tipo de cambio usado por Canavida desde el 10/10/2026
XLSX = "Copy of Canavida Gastos Ingresos Oficial (2).xlsx"
inv = json.load(open(sys.argv[1]))
hoy = sys.argv[2]

def safe(s): return re.sub(r"[^A-Za-z0-9._-]+", "-", s).strip("-")[:40]

# 1) PDFs por mes de emisión
for x in inv:
    x["neto"] = round(x["crc"] - x.get("ajuste_crc", 0), 2)
    x["usd"] = round(x["neto"] / TC, 2)
    d = f"facturas/{x['fecha'][:7]}"; os.makedirs(d, exist_ok=True)
    for i, p in enumerate(x.get("pdfs", [])):
        suf = "" if i == 0 else "_NC"
        nm = f"{x['fecha']}_{safe(x['prov'])}_{x['num'][-8:]}{suf}.pdf"
        shutil.copy(p, os.path.join(d, nm)); x.setdefault("files", []).append(f"{d}/{nm}")

# 2) Filas en Gastos Diarios (sin duplicar por número de factura)
wb = openpyxl.load_workbook(XLSX); ws = wb["Gastos Diarios"]
ya = {str(ws.cell(r, 7).value) for r in range(6, ws.max_row + 1)}
last = max(r for r in range(6, ws.max_row + 1) if ws.cell(r, 2).value is not None)
row = max(last + 1, 135)
filas = []
for x in sorted(inv, key=lambda v: (v["fecha"], v["prov"])):
    ref = f"Fact. {x['num'][-8:]}"
    if ref in ya: continue
    while ws.cell(row, 2).value is not None: row += 1
    vals = [dt.datetime.fromisoformat(x["fecha"]), x["cat"], x["desc"], x["prov"], x["metodo"],
            f"={x['neto']}/{TC}", ref, x.get("nota")]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row, c, v)
        if c == 1: cell.number_format = "dd/mm/yyyy"
        if c == 6: cell.number_format = '#,##0.00'
    x["fila"] = row; filas.append(row); row += 1
wb.save(XLSX)

# 3) Excel del día
os.makedirs("reportes", exist_ok=True)
h = ["Mes", "Fecha", "Tipo", "N°", "Proveedor", "Categoría", "Método", "Total ₡", "Nota crédito ₡", "Neto ₡", "USD (TC 459)", "Descripción", "Fila Gastos Diarios", "PDF"]
wb2 = openpyxl.Workbook(); s = wb2.active; s.title = "Facturas"; s.append(h)
for c in s[1]: c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="DDDDDD")
for x in inv:
    s.append([x["fecha"][:7], x["fecha"], x["tipo"], x["num"], x["prov"], x["cat"], x["metodo"], x["crc"],
              x.get("ajuste_crc") or None, x["neto"], x["usd"], x["desc"], x.get("fila"), ", ".join(x.get("files", []))])
for col, w in zip("ABCDEFGHIJKLMN", [9, 11, 10, 22, 30, 26, 20, 13, 13, 13, 12, 50, 10, 60]): s.column_dimensions[col].width = w
for r in s.iter_rows(min_row=2):
    for c in r[7:11]: c.number_format = '#,##0.00'
s2 = wb2.create_sheet("Resumen por mes"); s2.append(["Mes", "Categoría", "Facturas", "Neto ₡", "USD"])
for c in s2[1]: c.font = Font(bold=True)
agg = {}
for x in inv:
    a = agg.setdefault((x["fecha"][:7], x["cat"]), [0, 0, 0]); a[0] += 1; a[1] += x["neto"]; a[2] += x["usd"]
for (m, cat), (n, crc, usd) in sorted(agg.items()): s2.append([m, cat, n, round(crc, 2), round(usd, 2)])
tot_crc = round(sum(x["neto"] for x in inv), 2); tot_usd = round(sum(x["usd"] for x in inv), 2)
s2.append(["TOTAL", "", len(inv), tot_crc, tot_usd])
for c in s2[s2.max_row]: c.font = Font(bold=True)
for col, w in zip("ABCDE", [9, 30, 10, 15, 12]): s2.column_dimensions[col].width = w
for r in s2.iter_rows(min_row=2):
    for c in r[3:5]: c.number_format = '#,##0.00'
xl = f"reportes/facturas_{hoy}.xlsx"; wb2.save(xl)

# 4) PDF del día
st = getSampleStyleSheet()
for k in ("Normal", "Title", "Heading3"): st[k].fontName = "DVB" if k != "Normal" else "DV"
sm = st["Normal"].clone("sm"); sm.fontSize = 7; sm.leading = 8.5
fm = lambda v: f"{v:,.2f}"
d1 = [["Fecha", "N°", "Proveedor", "Categoría", "Método", "Neto ₡", "USD", "Descripción"]] + [
    [x["fecha"], x["num"][-8:], Paragraph(x["prov"], sm), Paragraph(x["cat"], sm), Paragraph(x["metodo"], sm),
     fm(x["neto"]), fm(x["usd"]), Paragraph(x["desc"] + (f" <i>({x['nota']})</i>" if x.get("nota") else ""), sm)] for x in inv]
d1.append(["", "", "TOTAL", "", "", fm(tot_crc), fm(tot_usd), ""])
t1 = Table(d1, repeatRows=1, colWidths=[55, 52, 120, 95, 70, 70, 55, 270])
t1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("GRID", (0, 0), (-1, -1), .3, colors.grey),
    ("FONTNAME", (0, 0), (-1, -1), "DV"), ("FONTSIZE", (0, 0), (-1, -1), 7), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (5, 1), (6, -1), "RIGHT"),
    ("FONTNAME", (0, -1), (-1, -1), "DVB")]))
d2 = [["Mes", "Categoría", "Facturas", "Neto ₡", "USD"]] + [[m, c, n, fm(a), fm(u)] for (m, c), (n, a, u) in sorted(agg.items())]
t2 = Table(d2, colWidths=[50, 150, 50, 80, 60])
t2.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey), ("GRID", (0, 0), (-1, -1), .3, colors.grey),
    ("FONTNAME", (0, 0), (-1, -1), "DV"), ("FONTSIZE", (0, 0), (-1, -1), 8), ("ALIGN", (2, 1), (-1, -1), "RIGHT")]))
fd = dt.date.fromisoformat(hoy).strftime("%d/%m/%Y")
pdf = f"reportes/resumen_facturas_{hoy}.pdf"
SimpleDocTemplate(pdf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24).build([
    Paragraph(f"Canavida · Facturas recibidas el {fd}", st["Title"]),
    Paragraph(f"{len(inv)} facturas · ₡{fm(tot_crc)} ≈ US${fm(tot_usd)} (TC ₡{TC}). PDF guardados por mes de emisión en facturas/AAAA-MM. "
              f"Filas agregadas a 'Gastos Diarios': {', '.join(map(str, filas)) or 'ninguna'}.", st["Normal"]),
    Spacer(1, 8), Paragraph("Resumen por mes y categoría", st["Heading3"]), t2, Spacer(1, 10),
    Paragraph("Detalle", st["Heading3"]), t1])
for x in inv: x.pop("pdfs", None)
json.dump(inv, open(f"reportes/facturas_{hoy}.json", "w"), ensure_ascii=False, indent=1)
print(xl, pdf, "filas", filas, "total", tot_crc, tot_usd)
