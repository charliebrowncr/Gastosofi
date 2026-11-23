#!/usr/bin/env python3
"""Genera la actualización diaria de facturas de Canavida.

Entrada:
  --resultados  carpeta con *.jsonl (una línea por PDF extraído del correo)
  --plantilla   Excel oficial de Gastos/Ingresos (formato del repositorio)
  --repo        raíz del repositorio (aquí se guardan facturas/AAAA-MM y el estado)
  --salida      carpeta donde se escriben el Excel y el PDF del día

Salida:
  facturas/AAAA-MM/<fecha>_<proveedor>_<numero>.pdf   (facturas acomodadas por mes)
  facturas/procesadas.json                             (estado: qué ya se reportó)
  <salida>/Canavida Gastos - Actualizado AAAA-MM-DD.xlsx
  <salida>/Canavida Facturas - Reporte AAAA-MM-DD.pdf
"""
import argparse, datetime as dt, glob, json, os, re, shutil, subprocess, unicodedata
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

TC = 459  # tipo de cambio ₡ por $ (el mismo que usa la hoja en noviembre)
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def slug(s, n=30):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-")
    return s[:n] or "sin-nombre"


def clave_unica(r):
    if r.get("clave"):
        return r["clave"]
    return f"{slug(r.get('proveedor'),40)}|{r.get('numero')}|{r.get('total')}"


def usd(r):
    if r.get("total") is None:
        return None
    return float(r["total"]) if r.get("moneda") == "USD" else float(r["total"]) / TC


def cargar(resultados):
    regs = {}
    for f in sorted(glob.glob(os.path.join(resultados, "*.jsonl"))):
        for line in open(f, encoding="utf-8"):
            line = line.strip()
            if not line:
                continue
            r = json.loads(line)
            if not r.get("es_factura"):
                continue
            k = clave_unica(r)
            if k not in regs:  # el mismo PDF puede llegar en varios correos
                regs[k] = r
    return regs


def eval_monto(v):
    if isinstance(v, (int, float)):
        return float(v)
    m = re.fullmatch(r"=\s*([\d.]+)\s*/\s*([\d.]+)", str(v or ""))
    return float(m.group(1)) / float(m.group(2)) if m else None


def ya_en_hoja(ws, r):
    """True si la factura ya fue anotada a mano en Gastos Diarios (mismo Nº y monto)."""
    try:
        num = int(re.sub(r"\D", "", str(r.get("numero") or ""))[-10:] or -1)
    except ValueError:
        return False
    u = usd(r)
    for row in range(6, ws.max_row + 1):
        g = ws.cell(row, 7).value
        try:
            g = int(g)
        except (TypeError, ValueError):
            continue
        if g == num:
            f = eval_monto(ws.cell(row, 6).value)
            if u is None or f is None or abs(f - u) <= max(1, 0.05 * u):
                return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--resultados", required=True)
    ap.add_argument("--plantilla", required=True)
    ap.add_argument("--repo", required=True)
    ap.add_argument("--salida", required=True)
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    a = ap.parse_args()

    hoy = a.fecha
    estado_f = os.path.join(a.repo, "facturas", "procesadas.json")
    estado = json.load(open(estado_f, encoding="utf-8")) if os.path.exists(estado_f) else {}
    regs = cargar(a.resultados)
    # incluir las ya procesadas en días anteriores (el estado guarda el registro completo)
    for k, r in estado.items():
        regs.setdefault(k, r)

    nuevas = []
    for k, r in regs.items():
        if k not in estado:
            r["reportada"] = hoy
            nuevas.append(k)
        mes = (r.get("fecha") or r.get("email_date", ""))[:7]
        destino_dir = os.path.join(a.repo, "facturas", mes)
        nombre = f"{r.get('fecha','')}_{slug(r.get('proveedor'))}_{slug(str(r.get('numero') or ''),25)}.pdf"
        destino = os.path.join(destino_dir, nombre)
        if not os.path.exists(destino) and r.get("pdf_path") and os.path.exists(r["pdf_path"]):
            os.makedirs(destino_dir, exist_ok=True)
            shutil.copyfile(r["pdf_path"], destino)
        r["archivo"] = os.path.relpath(destino, a.repo)
        estado[k] = r

    orden = sorted(regs.values(), key=lambda r: (r.get("fecha") or "", r.get("proveedor") or ""))
    por_mes = defaultdict(list)
    for r in orden:
        por_mes[(r.get("fecha") or "")[:7]].append(r)

    # ---------------- Excel ----------------
    wb = openpyxl.load_workbook(a.plantilla)
    ws = wb["Gastos Diarios"]
    fila = 6
    while ws.cell(fila, 2).value not in (None, "") or ws.cell(fila, 3).value not in (None, ""):
        fila += 1
    agregadas, duplicadas = [], []
    for r in orden:
        if r.get("total") is None:
            continue
        if ya_en_hoja(ws, r):
            duplicadas.append(r)
            r["en_hoja"] = "ya estaba"
            continue
        f = dt.datetime.strptime(r["fecha"], "%Y-%m-%d")
        ws.cell(fila, 1, f).number_format = "d/M/yyyy"
        ws.cell(fila, 2, r.get("categoria"))
        ws.cell(fila, 3, r.get("descripcion"))
        ws.cell(fila, 4, r.get("proveedor"))
        ws.cell(fila, 5, r.get("metodo_pago"))
        tot = float(r["total"])
        ws.cell(fila, 6, f"={round(tot, 2)}/{TC}" if r.get("moneda") == "CRC" else round(tot, 2))
        ws.cell(fila, 6).number_format = '\\$#,##0.00;"($"#,##0.00\\);\\-'
        num = re.sub(r"\D", "", str(r.get("numero") or ""))
        ws.cell(fila, 7, int(num[-10:]) if num else r.get("numero"))
        ws.cell(fila, 8, "Auto-correo" + (" · " + r["notas"] if r.get("notas") else ""))
        r["en_hoja"] = f"fila {fila}"
        agregadas.append(r)
        fila += 1

    hj = wb.create_sheet("Facturas por Mes")
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="1F4E78")
    mes_fill = PatternFill("solid", fgColor="DDEBF7")
    hj["A1"] = f"Facturas PDF recibidas por correo · actualizado {hoy}"
    hj["A1"].font = Font(bold=True, size=14)
    hj["A2"] = f"Montos en colones convertidos a USD con ₡{TC} por $. 'Nueva' = reportada hoy."
    cols = ["Fecha", "Proveedor", "Nº factura", "Descripción", "Categoría", "Método de pago",
            "Moneda", "Total (moneda original)", "Monto (USD)", "Nueva", "En Gastos Diarios", "Archivo PDF", "Notas"]
    row = 4
    for mes in sorted(por_mes):
        y, m = (int(x) for x in mes.split("-"))
        hj.cell(row, 1, f"{MESES[m]} {y}").font = Font(bold=True, size=12)
        for c in range(1, len(cols) + 1):
            hj.cell(row, c).fill = mes_fill
        row += 1
        for i, h in enumerate(cols, 1):
            c = hj.cell(row, i, h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
        row += 1
        ini = row
        for r in por_mes[mes]:
            vals = [dt.datetime.strptime(r["fecha"], "%Y-%m-%d"), r.get("proveedor"), str(r.get("numero") or ""),
                    r.get("descripcion"), r.get("categoria"), r.get("metodo_pago"), r.get("moneda"),
                    r.get("total"), round(usd(r), 2) if usd(r) is not None else None,
                    "Sí" if r.get("reportada") == hoy else "", r.get("en_hoja", ""), r.get("archivo"), r.get("notas", "")]
            for i, v in enumerate(vals, 1):
                hj.cell(row, i, v)
            hj.cell(row, 1).number_format = "d/M/yyyy"
            hj.cell(row, 8).number_format = "#,##0.00"
            hj.cell(row, 9).number_format = "$#,##0.00"
            row += 1
        hj.cell(row, 8, "Total mes").font = bold
        hj.cell(row, 9, f"=SUM(I{ini}:I{row-1})").font = bold
        hj.cell(row, 9).number_format = "$#,##0.00"
        row += 2
    for i, w in enumerate([11, 28, 22, 40, 28, 20, 8, 16, 13, 7, 16, 60, 40], 1):
        hj.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    hj.freeze_panes = "A4"

    os.makedirs(a.salida, exist_ok=True)
    xlsx = os.path.join(a.salida, f"Canavida Gastos - Actualizado {hoy}.xlsx")
    wb.save(xlsx)

    # ---------------- PDF ----------------
    pdf = os.path.join(a.salida, f"Canavida Facturas - Reporte {hoy}.pdf")
    st = getSampleStyleSheet()
    small = st["BodyText"].clone("small", fontSize=7.5, leading=9)
    doc = SimpleDocTemplate(pdf, pagesize=landscape(letter), leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch, title=f"Facturas Canavida {hoy}")
    el = [Paragraph("Canavida Villas &amp; Resort – Facturas recibidas por correo", st["Title"]),
          Paragraph(f"Actualización del {hoy} · {len(regs)} facturas en total · {len(nuevas)} nuevas hoy · "
                    f"Colones convertidos a USD con ₡{TC}/$", st["BodyText"]), Spacer(1, 8)]
    ts = TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                     ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTSIZE", (0, 0), (-1, -1), 8),
                     ("GRID", (0, 0), (-1, -1), 0.25, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                     ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")])])
    # resumen por mes y categoría
    cats = sorted({r.get("categoria") or "Sin categoría" for r in orden})
    meses = sorted(por_mes)
    data = [["Categoría"] + [f"{MESES[int(m[5:])]} {m[:4]}" for m in meses] + ["Total"]]
    for c in cats:
        fila_v = [sum(usd(r) or 0 for r in por_mes[m] if (r.get("categoria") or "Sin categoría") == c) for m in meses]
        data.append([c] + [f"${v:,.2f}" for v in fila_v] + [f"${sum(fila_v):,.2f}"])
    tot = [sum(usd(r) or 0 for r in por_mes[m]) for m in meses]
    data.append(["TOTAL"] + [f"${v:,.2f}" for v in tot] + [f"${sum(tot):,.2f}"])
    data.append(["Nº de facturas"] + [str(len(por_mes[m])) for m in meses] + [str(len(orden))])
    t = Table(data, repeatRows=1)
    t.setStyle(ts)
    t.setStyle(TableStyle([("FONTNAME", (0, -2), (-1, -1), "Helvetica-Bold"), ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
    el += [Paragraph("Resumen por mes y categoría (USD)", st["Heading2"]), t, Spacer(1, 10)]

    def tabla(rs):
        d = [["Fecha", "Proveedor", "Nº factura", "Descripción", "Categoría", "Total", "USD", "Nueva"]]
        for r in rs:
            tot_txt = "" if r.get("total") is None else (
                f"₡{float(r['total']):,.2f}" if r.get("moneda") == "CRC" else f"${float(r['total']):,.2f}")
            u = usd(r)
            d.append([r.get("fecha"), Paragraph(str(r.get("proveedor") or ""), small),
                      Paragraph(str(r.get("numero") or ""), small), Paragraph(str(r.get("descripcion") or ""), small),
                      Paragraph(str(r.get("categoria") or ""), small), tot_txt,
                      "" if u is None else f"${u:,.2f}", "Sí" if r.get("reportada") == hoy else ""])
        tb = Table(d, repeatRows=1, colWidths=[0.75 * inch, 1.9 * inch, 1.6 * inch, 2.6 * inch, 1.5 * inch,
                                                0.95 * inch, 0.75 * inch, 0.45 * inch])
        tb.setStyle(ts)
        tb.setStyle(TableStyle([("ALIGN", (5, 1), (6, -1), "RIGHT")]))
        return tb

    nuevas_r = [r for r in orden if r.get("reportada") == hoy]
    if nuevas_r:
        el += [Paragraph(f"Facturas nuevas de hoy ({len(nuevas_r)})", st["Heading2"]), tabla(nuevas_r)]
    sin_monto = [r for r in orden if r.get("total") is None]
    if sin_monto:
        el += [Spacer(1, 8), Paragraph("Facturas que requieren revisión manual (no se pudo leer el total)", st["Heading3"]),
               tabla(sin_monto)]
    for m in meses:
        el += [PageBreak(), Paragraph(f"{MESES[int(m[5:])]} {m[:4]} – {len(por_mes[m])} facturas · "
                                      f"${sum(usd(r) or 0 for r in por_mes[m]):,.2f}", st["Heading2"]), tabla(por_mes[m])]
    doc.build(el)

    # Anexar las facturas nuevas al PDF si el tamaño lo permite (límite de correo 25 MB)
    anexos = [os.path.join(a.repo, r["archivo"]) for r in nuevas_r if os.path.exists(os.path.join(a.repo, r["archivo"]))]
    if anexos and sum(os.path.getsize(p) for p in anexos) < 15 * 1024 * 1024:
        tmp = pdf + ".tmp.pdf"
        res = subprocess.run(["pdfunite", pdf] + anexos + [tmp], capture_output=True)
        if res.returncode == 0:
            os.replace(tmp, pdf)

    with open(estado_f, "w", encoding="utf-8") as fh:
        json.dump(estado, fh, ensure_ascii=False, indent=1, sort_keys=True)

    resumen = {"total_facturas": len(regs), "nuevas_hoy": len(nuevas), "filas_agregadas_excel": len(agregadas),
               "ya_estaban_en_hoja": len(duplicadas), "sin_monto": len(sin_monto),
               "por_mes_usd": {m: round(sum(usd(r) or 0 for r in por_mes[m]), 2) for m in meses},
               "por_mes_n": {m: len(por_mes[m]) for m in meses},
               "xlsx": xlsx, "pdf": pdf, "pdf_bytes": os.path.getsize(pdf)}
    print(json.dumps(resumen, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
