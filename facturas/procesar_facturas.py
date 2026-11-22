#!/usr/bin/env python3
"""Procesa facturas descargadas de Gmail y genera el Excel (formato Canavida) y un reporte PDF.

Entrada: un directorio "inbox" con una carpeta por mensaje de Gmail:
    inbox/<messageId>/meta.json   (messageId, date, subject, sender, bodyExcerpt, files)
    inbox/<messageId>/*.pdf       (factura en PDF)
    inbox/<messageId>/*.xml       (comprobante electrónico de Hacienda, opcional)

Estado: el registro acumulado vive en la hoja oculta "_registro" del Excel generado.
Cada corrida lee el Excel anterior (--previo), agrega solo facturas nuevas (por clave
o por messageId+archivo) y vuelve a construir el libro desde la plantilla.

Uso:
  python3 procesar_facturas.py --plantilla PLANTILLA.xlsx --inbox DIR --salida DIR \
      [--previo ANTERIOR.xlsx] [--tc 459] [--fecha 2026-11-22]
"""
import argparse
import hashlib
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import unicodedata
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]
CAMPOS = ["clave", "messageId", "archivo", "fecha", "proveedor", "cedula", "numero",
          "descripcion", "categoria", "metodo", "moneda", "monto_original", "tc",
          "monto_usd", "tipo", "estado", "notas", "fecha_correo", "asunto", "procesado", "sha"]

CATEGORIAS_KW = [
    ("Servicios públicos e internet", ["ice", "instituto costarricense de electricidad", "kolbi", "telefon", "internet", "cnfl", "acueducto", "asada", "coopeguanacaste", "coopealfaro"]),
    ("Combustible y transporte", ["gasolin", "servicentro", "combustible", "coopeagri", "estacion de servicio", "gs ", "uno ", "delta", "taxi"]),
    ("Mantenimiento y reparaciones", ["ferreter", "material", "construc", "coopemadereros", "electric", "colono", "epa", "repuesto", "taller", "plomer", "pintura", "sur quimica"]),
    ("Alimentos y bebidas", ["supermerc", "super ", "automercado", "mas x menos", "maxi pali", "pali", "walmart", "fruta", "verdura", "carniceria", "panaderia", "restaurante", "rest.", "bm de uvita", "bm supermercado", "soda", "licorera", "spirits"]),
    ("Limpieza y lavandería", ["lavander", "limpieza", "cleaner", "natural center"]),
    ("Suministros y amenities", ["pequeño mundo", "pequeno mundo", "ekono", "libreria", "bazar", "gas", "amenit"]),
    ("Jardinería y piscinas", ["piscina", "jardin", "vivero", "agro", "cloro"]),
    ("Software y tecnología", ["software", "smart506", "pricelabs", "breezeway", "wireless", "computo", "tecnolog", "quo"]),
    ("Comisiones de plataformas", ["booking", "airbnb", "expedia", "lodgify"]),
    ("Seguros", ["seguro", "ins ", "instituto nacional de seguros"]),
    ("Impuestos y patentes", ["municipalidad", "hacienda", "patente", "impuesto"]),
    ("Cargas sociales (CCSS/INS)", ["ccss", "caja costarricense"]),
    ("Imprevistos y otros", ["contador", "accounting", "abogad", "medic", "farmacia", "urgent care", "clinica"]),
    ("Marketing y publicidad", ["marketing", "publicidad", "facebook", "meta platforms", "google ads"]),
]
METODO_PAGO = {"01": "Efectivo", "02": "Tarjeta", "03": "Transferencia / SINPE",
               "04": "Transferencia / SINPE", "06": "Transferencia / SINPE"}


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def local(tag):
    return tag.rsplit("}", 1)[-1]


def find(el, *path):
    """Busca por nombre local ignorando namespaces."""
    cur = [el]
    for name in path:
        nxt = []
        for c in cur:
            nxt.extend(ch for ch in c.iter() if local(ch.tag) == name)
        cur = nxt
        if not cur:
            return None
    return cur[0]


def text(el, *path):
    e = find(el, *path)
    return e.text.strip() if e is not None and e.text else None


def parse_xml(path):
    try:
        root = ET.parse(path).getroot()
    except Exception:
        return None
    tipo_doc = local(root.tag)
    if tipo_doc not in ("FacturaElectronica", "TiqueteElectronico", "NotaCreditoElectronica",
                        "NotaDebitoElectronica", "FacturaElectronicaCompra", "FacturaElectronicaExportacion"):
        return None
    emisor = find(root, "Emisor")
    receptor = find(root, "Receptor")
    resumen = find(root, "ResumenFactura")
    lineas = [ch for ch in root.iter() if local(ch.tag) == "LineaDetalle"]
    detalles = [text(l, "Detalle") for l in lineas if text(l, "Detalle")]
    desc = "; ".join(detalles[:3]) + (f" (+{len(detalles) - 3} más)" if len(detalles) > 3 else "")
    medio = None
    for ch in root.iter():
        if local(ch.tag) in ("MedioPago", "TipoMedioPago") and ch.text and ch.text.strip():
            medio = ch.text.strip()
            break
    total = text(resumen, "TotalComprobante") if resumen is not None else None
    fecha = text(root, "FechaEmision")
    return {
        "tipo_doc": tipo_doc,
        "clave": text(root, "Clave"),
        "numero": text(root, "NumeroConsecutivo"),
        "fecha": fecha[:10] if fecha else None,
        "proveedor": text(emisor, "Nombre") if emisor is not None else None,
        "comercial": text(emisor, "NombreComercial") if emisor is not None else None,
        "cedula": text(emisor, "Numero") if emisor is not None else None,
        "receptor": text(receptor, "Nombre") if receptor is not None else None,
        "moneda": (text(resumen, "CodigoMoneda") if resumen is not None else None) or "CRC",
        "tc": float(text(resumen, "TipoCambio") or 1) if resumen is not None else 1.0,
        "total": round(float(total), 2) if total else None,
        "descripcion": desc,
        "metodo": METODO_PAGO.get(medio, ""),
    }


def pdf_text(path):
    try:
        out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, text=True, timeout=60)
        return out.stdout
    except Exception:
        return ""


NUM = r"([0-9]{1,3}(?:[.,][0-9]{3})*(?:[.,][0-9]{1,2})|[0-9]+(?:[.,][0-9]{1,2})?)"


def to_float(s):
    s = s.strip()
    if re.search(r",\d{1,2}$", s) and "." in s:  # 1.234,56
        s = s.replace(".", "").replace(",", ".")
    elif re.search(r",\d{1,2}$", s):
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_pdf(path, meta):
    t = pdf_text(path)
    info = {"texto": t}
    m = re.search(r"\b(506\d{47})\b", t.replace(" ", "")) or re.search(r"\b(506\d{47})\b", meta.get("subject", "") + " " + os.path.basename(path))
    if m:
        info["clave"] = m.group(1)
        info["numero"] = m.group(1)[21:41]
        c = m.group(1)
        info["fecha"] = f"20{c[7:9]}-{c[5:7]}-{c[3:5]}"
        info["cedula"] = c[9:21].lstrip("0")
    if "fecha" not in info:
        m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](20\d{2})", t)
        if m:
            info["fecha"] = f"{m.group(3)}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    totals = []
    for line in t.splitlines():
        l = norm(line)
        if re.search(r"total\s*(comprobante|factura|a pagar|general|due|amount)?\b", l) and not re.search(r"sub\s*total|total\s*(impuesto|iva|descuento|gravad|exent|exoner|servicio|mercanc|venta)", l):
            nums = [to_float(x) for x in re.findall(NUM, line)]
            nums = [n for n in nums if n]
            if nums:
                totals.append(nums[-1])
    if totals:
        info["total"] = totals[-1]
    tl = t.lower()
    colones = "₡" in t or re.search(r"\bcrc\b|colones", tl)
    dolares = "$" in t or re.search(r"\busd\b|d[oó]lar", tl)
    info["moneda"] = "USD" if dolares and not colones else "CRC"
    first = next((ln.strip() for ln in t.splitlines() if ln.strip()), "")
    info["proveedor"] = re.split(r"\s{2,}|\|", first)[0][:80] if first else None
    return info


def guess_proveedor_asunto(subject):
    s = re.sub(r"^(rv|re|fw|fwd)\s*:\s*", "", subject or "", flags=re.I)
    for pat in [r"Emisor:\s*([^,\n]+?)(?:\s+Notificacion|$|,)", r"emitida por\s+(.+?)\s+para", r"de\s+(.+?)\s+-", r"-\s*([^-]+)$"]:
        m = re.search(pat, s, flags=re.I)
        if m:
            return m.group(1).strip()
    return s.strip()[:80]


def historial_proveedores(ws):
    hist = defaultdict(Counter)
    for r in ws.iter_rows(min_row=6, values_only=True):
        if r[1] and r[3]:
            hist[norm(r[3])][r[1]] += 1
    return hist


def categorizar(proveedor, descripcion, hist):
    p = norm(proveedor)
    for k, c in hist.items():
        if k and len(k) >= 3 and (k in p or p.startswith(k) or (p and p in k)):
            return c.most_common(1)[0][0]
    blob = f" {p} {norm(descripcion)} "
    for cat, kws in CATEGORIAS_KW:
        if any(kw in blob for kw in kws):
            return cat
    return "Imprevistos y otros"


def facturas_manuales(ws):
    """(fecha, numero) registrados a mano en Gastos Diarios para detectar duplicados."""
    s = set()
    for r in ws.iter_rows(min_row=6, values_only=True):
        if r[0] and r[6] not in (None, ""):
            try:
                s.add(str(int(float(r[6]))))
            except (TypeError, ValueError):
                s.add(str(r[6]).strip())
    return s


def es_emitida_por_canavida(nombre):
    n = norm(nombre)
    return "canavida" in n


def leer_registro(path):
    if not path or not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path)
    if "_registro" not in wb.sheetnames:
        return []
    ws = wb["_registro"]
    rows = list(ws.iter_rows(values_only=True))
    head = rows[0]
    return [dict(zip(head, r)) for r in rows[1:] if any(v is not None for v in r)]


def procesar_inbox(inbox, registro, hist, manuales, tc_default, hoy):
    vistos = {r["clave"] for r in registro if r.get("clave")} | {f'{r["messageId"]}/{r["archivo"]}' for r in registro}
    vistos |= {r["sha"] for r in registro if r.get("sha")}
    nuevos = []
    for mid in sorted(os.listdir(inbox)):
        d = os.path.join(inbox, mid)
        mp = os.path.join(d, "meta.json")
        if not os.path.isfile(mp):
            continue
        meta = json.load(open(mp, encoding="utf-8"))
        files = os.listdir(d)
        pdfs = sorted(f for f in files if f.lower().endswith(".pdf"))
        xmls = [parse_xml(os.path.join(d, f)) for f in files if f.lower().endswith(".xml")]
        xmls = [x for x in xmls if x]
        for pdf in pdfs:
            sha = hashlib.sha256(open(os.path.join(d, pdf), "rb").read()).hexdigest()
            if sha in vistos:
                continue
            pinfo = parse_pdf(os.path.join(d, pdf), meta)
            # Asociar XML por clave (o el único XML del correo)
            x = None
            for cand in xmls:
                if cand["clave"] and (cand["clave"] in pdf or cand["clave"] == pinfo.get("clave")):
                    x = cand
            if x is None and len(xmls) == 1 and len(pdfs) == 1:
                x = xmls[0]
            clave = (x or {}).get("clave") or pinfo.get("clave")
            key = f"{mid}/{pdf}"
            if (clave and clave in vistos) or key in vistos:
                continue
            fecha = (x or {}).get("fecha") or pinfo.get("fecha") or (meta.get("date") or "")[:10]
            proveedor = (x or {}).get("comercial") or (x or {}).get("proveedor") or guess_proveedor_asunto(meta.get("subject")) or pinfo.get("proveedor")
            moneda = (x or {}).get("moneda") or pinfo.get("moneda") or "CRC"
            total = (x or {}).get("total") if x else pinfo.get("total")
            if moneda == "USD":
                tc, usd = 1.0, total
            else:
                tc = tc_default
                usd = round(total / tc, 2) if total else None
            descripcion = (x or {}).get("descripcion") or re.sub(r"^(rv|re|fw|fwd)\s*:\s*", "", meta.get("subject") or "", flags=re.I)[:120]
            numero = (x or {}).get("numero") or pinfo.get("numero") or ""
            subj = norm(meta.get("subject"))
            if es_emitida_por_canavida((x or {}).get("proveedor") or "") or (not x and re.search(r"factura no\.\d+ - (cr )?canavida", subj)):
                tipo = "Emitida por Canavida (ingreso)"
            elif "pre-factura" in subj or "proforma" in subj or "cotizacion" in subj:
                tipo = "Pre-factura / cotización"
            elif (x or {}).get("tipo_doc") == "NotaCreditoElectronica":
                tipo = "Nota de crédito"
            else:
                tipo = "Gasto"
            notas = []
            corto = str(int(numero[-10:])) if numero and numero[-10:].isdigit() else None
            estado = "Nuevo"
            if corto and corto in manuales:
                estado = "Posible duplicado (ya en Gastos Diarios)"
            if total is None:
                estado = "Revisar monto"
                notas.append("No se pudo leer el total del PDF")
            if not x:
                notas.append("Datos leídos del PDF (sin XML)")
            nuevos.append({
                "clave": clave, "messageId": mid, "archivo": pdf, "fecha": fecha,
                "proveedor": proveedor, "cedula": (x or {}).get("cedula") or pinfo.get("cedula"),
                "numero": numero, "descripcion": descripcion,
                "categoria": categorizar(proveedor, descripcion, hist) if tipo == "Gasto" else "",
                "metodo": (x or {}).get("metodo", ""), "moneda": moneda, "monto_original": total,
                "tc": tc, "monto_usd": usd, "tipo": tipo, "estado": estado, "notas": "; ".join(notas),
                "fecha_correo": (meta.get("date") or "")[:10], "asunto": meta.get("subject"),
                "procesado": hoy,
            })
            nuevos[-1]["sha"] = sha
            vistos.add(sha)
            if clave:
                vistos.add(clave)
            vistos.add(key)
    return nuevos


def nombre_archivo(r):
    prov = re.sub(r"[^A-Za-z0-9]+", "_", norm(r["proveedor"]))[:40].strip("_") or "proveedor"
    num = (r.get("numero") or "")[-10:] or r["messageId"][:8]
    return f'{r["fecha"]}_{prov}_{num}.pdf'


def mes_de(fecha):
    y, m = int(fecha[:4]), int(fecha[5:7])
    return f"{y}-{m:02d} {MESES[m]}", f"{MESES[m]} {y}"


AZUL = PatternFill("solid", fgColor="DDEBF7")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
ROJO = PatternFill("solid", fgColor="F8CBAD")
HEAD = PatternFill("solid", fgColor="1F4E78")


def construir_excel(plantilla, registro, salida_xlsx, hoy):
    wb = openpyxl.load_workbook(plantilla)
    gd = wb["Gastos Diarios"]
    # Agregar a "Gastos Diarios" solo gastos nuevos (no duplicados) dentro del periodo de la plantilla
    ini, fin = dt.date(2026, 10, 1), dt.date(2026, 12, 31)
    last = 5
    for row in range(6, gd.max_row + 1):
        if any(gd.cell(row, c).value not in (None, "") for c in range(2, 8)):
            last = row
    fila = last + 1
    for r in sorted(registro, key=lambda r: r["fecha"]):
        f = dt.date.fromisoformat(r["fecha"])
        if r["tipo"] != "Gasto" or r["estado"] != "Nuevo" or not (ini <= f <= fin):
            continue
        monto = r["monto_usd"]
        if r["moneda"] != "USD" and r["monto_original"]:
            monto = f'={r["monto_original"]}/{int(r["tc"]) if float(r["tc"]).is_integer() else r["tc"]}'
        vals = [dt.datetime.combine(f, dt.time()), r["categoria"], (r["descripcion"] or "")[:120], r["proveedor"],
                r["metodo"] or None, monto, (r["numero"] or "")[-10:].lstrip("0") or None, "Auto: Gmail"]
        for c, v in enumerate(vals, 1):
            gd.cell(fila, c, v)
        gd.cell(fila, 1).number_format = "dd/mm/yyyy"
        gd.cell(fila, 6).number_format = '#,##0.00'
        fila += 1

    # Una hoja por mes con todas las facturas PDF
    por_mes = defaultdict(list)
    for r in registro:
        por_mes[r["fecha"][:7]].append(r)
    cols = ["Fecha", "Categoría", "Descripción", "Proveedor", "Método de pago", "Monto (USD)",
            "Nº factura / Notas", "Moneda", "Monto original", "Tipo de cambio", "Tipo", "Estado",
            "Archivo PDF", "Notas"]
    resumen = defaultdict(lambda: defaultdict(float))
    for ym in sorted(por_mes):
        _, titulo = mes_de(ym + "-01")
        ws = wb.create_sheet(f"Facturas {titulo}"[:31])
        ws["A1"] = f"Facturas PDF recibidas por correo – {titulo}"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A2"] = f"Generado automáticamente el {hoy}. Monto USD = monto en colones / tipo de cambio (edite la columna J si hace falta)."
        for c, h in enumerate(cols, 1):
            cell = ws.cell(4, c, h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = HEAD
            cell.alignment = Alignment(wrap_text=True, vertical="center")
        rows = sorted(por_mes[ym], key=lambda r: (r["fecha"], r["proveedor"] or ""))
        for i, r in enumerate(rows, 5):
            f = dt.date.fromisoformat(r["fecha"])
            ws.cell(i, 1, dt.datetime.combine(f, dt.time())).number_format = "dd/mm/yyyy"
            ws.cell(i, 2, r["categoria"])
            ws.cell(i, 3, r["descripcion"])
            ws.cell(i, 4, r["proveedor"])
            ws.cell(i, 5, r["metodo"])
            if r["monto_original"] is not None:
                ws.cell(i, 6, f"=IF(N(J{i})=0,\"\",I{i}/J{i})").number_format = '#,##0.00'
            ws.cell(i, 7, r["numero"])
            ws.cell(i, 8, r["moneda"])
            ws.cell(i, 9, r["monto_original"]).number_format = '#,##0.00'
            ws.cell(i, 10, r["tc"])
            ws.cell(i, 11, r["tipo"])
            ws.cell(i, 12, r["estado"])
            ws.cell(i, 13, nombre_archivo(r))
            ws.cell(i, 14, r["notas"])
            fill = ROJO if r["estado"] == "Revisar monto" else AMARILLO if r["estado"].startswith("Posible") else None
            if fill:
                for c in range(1, len(cols) + 1):
                    ws.cell(i, c).fill = fill
            if r["tipo"] == "Gasto" and r["monto_usd"]:
                resumen[ym][r["categoria"]] += r["monto_usd"]
        n = len(rows) + 4
        ws.cell(n + 2, 5, "TOTAL GASTOS (USD)").font = Font(bold=True)
        ws.cell(n + 2, 6, f'=SUMIFS(F5:F{n},K5:K{n},"Gasto")').number_format = '#,##0.00'
        ws.cell(n + 2, 6).font = Font(bold=True)
        for c, w in enumerate([12, 26, 50, 32, 20, 14, 22, 8, 16, 10, 26, 30, 50, 34], 1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{n}"

    # Hoja oculta con el registro completo (estado para la próxima corrida)
    reg = wb.create_sheet("_registro")
    reg.append(CAMPOS)
    for r in registro:
        reg.append([r.get(k) for k in CAMPOS])
    reg.sheet_state = "hidden"
    wb.save(salida_xlsx)
    return resumen


def construir_pdf(registro, nuevos, resumen, salida_pdf, hoy):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    st = getSampleStyleSheet()
    small = st["BodyText"].clone("small", fontSize=7.5, leading=9)
    doc = SimpleDocTemplate(salida_pdf, pagesize=landscape(letter), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm, title=f"Facturas Canavida {hoy}")
    el = [Paragraph(f"Canavida – Reporte diario de facturas ({hoy})", st["Title"])]
    gastos = [r for r in registro if r["tipo"] == "Gasto"]
    el.append(Paragraph(
        f"Facturas PDF registradas: <b>{len(registro)}</b> · nuevas hoy: <b>{len(nuevos)}</b> · "
        f"por revisar: <b>{sum(r['estado'] != 'Nuevo' for r in registro)}</b>", st["Normal"]))
    el.append(Spacer(1, 8))

    def tabla(data, widths, header_bg="#1F4E78"):
        t = Table(data, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
        ]))
        return t

    # Resumen por mes y categoría
    meses = sorted(resumen)
    cats = sorted({c for m in meses for c in resumen[m]})
    if meses:
        el.append(Paragraph("Gastos por mes y categoría (USD, solo facturas tipo Gasto)", st["Heading2"]))
        data = [["Categoría"] + [mes_de(m + "-01")[1] for m in meses] + ["Total"]]
        for c in cats:
            vals = [resumen[m].get(c, 0) for m in meses]
            data.append([c] + [f"{v:,.2f}" for v in vals] + [f"{sum(vals):,.2f}"])
        tot = [sum(resumen[m].values()) for m in meses]
        data.append(["TOTAL"] + [f"{v:,.2f}" for v in tot] + [f"{sum(tot):,.2f}"])
        t = tabla(data, [6 * cm] + [3 * cm] * (len(meses) + 1))
        t.setStyle(TableStyle([("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
        el.append(t)

    def lista(titulo, rows):
        el.append(Paragraph(titulo, st["Heading2"]))
        if not rows:
            el.append(Paragraph("Ninguna.", st["Normal"]))
            return
        data = [["Fecha", "Proveedor", "Nº factura", "Categoría / Tipo", "Moneda", "Monto", "USD", "Estado"]]
        for r in sorted(rows, key=lambda r: r["fecha"]):
            data.append([r["fecha"], Paragraph(str(r["proveedor"] or ""), small), (r["numero"] or "")[-10:],
                         Paragraph(r["categoria"] or r["tipo"], small), r["moneda"],
                         f'{r["monto_original"]:,.2f}' if r["monto_original"] is not None else "?",
                         f'{r["monto_usd"]:,.2f}' if r["monto_usd"] is not None else "?",
                         Paragraph(r["estado"] + (f" – {r['notas']}" if r["notas"] else ""), small)])
        el.append(tabla(data, [2.1 * cm, 6 * cm, 2.6 * cm, 4.2 * cm, 1.4 * cm, 2.6 * cm, 2.1 * cm, 4.6 * cm]))

    lista(f"Facturas nuevas en esta actualización ({len(nuevos)})", nuevos)
    el.append(PageBreak())
    for m in sorted({r["fecha"][:7] for r in registro}):
        lista(f"Todas las facturas – {mes_de(m + '-01')[1]}", [r for r in registro if r["fecha"][:7] == m])
    doc.build(el)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plantilla", required=True)
    ap.add_argument("--inbox", required=True)
    ap.add_argument("--salida", required=True)
    ap.add_argument("--previo")
    ap.add_argument("--tc", type=float, default=459.0, help="Tipo de cambio ₡ por $ para facturas en colones")
    ap.add_argument("--fecha", default=dt.date.today().isoformat())
    a = ap.parse_args()

    os.makedirs(a.salida, exist_ok=True)
    plant = openpyxl.load_workbook(a.plantilla)
    hist = historial_proveedores(plant["Gastos Diarios"])
    manuales = facturas_manuales(plant["Gastos Diarios"])
    registro = leer_registro(a.previo)
    nuevos = procesar_inbox(a.inbox, registro, hist, manuales, a.tc, a.fecha)
    registro = registro + nuevos

    # PDFs organizados por mes
    pdf_dir = os.path.join(a.salida, "pdf")
    for r in nuevos:
        carpeta, _ = mes_de(r["fecha"])
        dest = os.path.join(pdf_dir, carpeta)
        os.makedirs(dest, exist_ok=True)
        shutil.copy2(os.path.join(a.inbox, r["messageId"], r["archivo"]), os.path.join(dest, nombre_archivo(r)))

    xlsx = os.path.join(a.salida, f"Canavida Facturas {a.fecha}.xlsx")
    pdf = os.path.join(a.salida, f"Canavida Reporte Facturas {a.fecha}.pdf")
    resumen = construir_excel(a.plantilla, registro, xlsx, a.fecha)
    construir_pdf(registro, nuevos, resumen, pdf, a.fecha)
    json.dump({"nuevos": nuevos, "total_registro": len(registro), "xlsx": xlsx, "pdf": pdf},
              open(os.path.join(a.salida, "resultado.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print(f"registro={len(registro)} nuevos={len(nuevos)} xlsx={xlsx} pdf={pdf}")


if __name__ == "__main__":
    main()
