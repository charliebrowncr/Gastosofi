#!/usr/bin/env python3
"""Procesa facturas descargadas de Gmail y genera el reporte diario.

Entrada : _inbox/<message_id>/{meta.json, *.pdf, *.xml}
Salida  : facturas/AAAA-MM/<fecha>_<proveedor>_<consecutivo>.pdf   (archivo por mes)
          data/facturas.csv                                        (registro maestro, idempotente)
          reportes/Canavida_Gastos_Ingresos_<fecha>.xlsx           (plantilla + facturas en "Gastos Diarios")
          reportes/Reporte_Facturas_<fecha>.pdf                    (resumen por mes)
"""
import csv, json, re, shutil, sys, unicodedata
from collections import defaultdict
from datetime import date, datetime
from copy import copy
from pathlib import Path
import xml.etree.ElementTree as ET

import openpyxl
import pdfplumber

ROOT = Path(__file__).resolve().parent.parent
INBOX = ROOT / "_inbox"
FACT_DIR = ROOT / "facturas"
LEDGER = ROOT / "data" / "facturas.csv"
REP_DIR = ROOT / "reportes"
TEMPLATE = ROOT / "Copy of Canavida Gastos Ingresos Oficial (2).xlsx"
CANAVIDA_ID = "3102777726"
TIPO_CAMBIO = 459.0  # ₡ por $ (el mismo que usa la hoja desde mediados de octubre)
FECHA_MIN, FECHA_MAX = date(2026, 10, 1), date(2026, 12, 31)
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
         "Septiembre", "Octubre", "Noviembre", "Diciembre"]

FIELDS = ["message_id", "clave", "consecutivo", "tipo_doc", "fecha", "mes", "proveedor", "nombre_comercial",
          "cedula_emisor", "descripcion", "categoria", "metodo_pago", "moneda", "total_moneda",
          "tipo_cambio", "monto_usd", "pdf", "fuente", "estado", "notas", "procesado"]

# palabra clave (en proveedor / nombre comercial / detalle) -> categoría de la hoja Listas
REGLAS = [
    (r"supermercado|maxi ?pal|pali|walmart|bm |\bbm\b|mas x menos|frutas|verdura|carnice|panader|abastecedor|mini ?super", "Alimentos y bebidas"),
    (r"restaur|rest\.|soda|burrito|pizz|cafe|caf[eé]|bar |hornito|mariscos|fogata|sushi|taco", "Alimentos y bebidas"),
    (r"ferreter|material|colono|madera|coopemadereros|construc|el[eé]ctric|plomer|pintur|sur quimica", "Mantenimiento y reparaciones"),
    (r"\bice\b|instituto costarricense de electricidad|kolbi|tel[eé]fono|\btel:|internet|wireless|cable|liberty|claro|movistar|coopeagri.*electric|cnfl|acueducto|\baya\b", "Servicios públicos e internet"),
    (r"gasolin|servicentro|combustible|estaci[oó]n de servicio|\bgs\b|delta|uno ", "Combustible y transporte"),
    (r"jard[ií]n|piscina|vivero|botica|agro", "Jardinería y piscinas"),
    (r"lavander|limpieza|cleaner|clean", "Limpieza y lavandería"),
    (r"amenit|suministro|libreri|bazar|ekono|oficina", "Suministros y amenities"),
    (r"software|tecnolog|smart506|pricelabs|breezeway|lodgify|google|microsoft|apple|computo", "Software y tecnología"),
    (r"mueble|muebler", "Muebleria"),
    (r"seguro|\bins\b|aseguradora", "Seguros"),
    (r"municipal|patente|hacienda|tributa", "Impuestos y patentes"),
    (r"publicidad|marketing|facebook|meta platforms", "Marketing y publicidad"),
    (r"booking|airbnb|expedia|comisi", "Comisiones de plataformas"),
    (r"tour|actividad", "Costo de tours y actividades"),
    (r"ccss|caja costarricense", "Cargas sociales (CCSS/INS)"),
    (r"contador|accounting|contab|abogad|legal|m[eé]dic|urgent|farmac|repuesto|taller", "Imprevistos y otros"),
]
MEDIO_PAGO = {"01": "Efectivo", "02": "Tarjeta", "03": "Transferencia / SINPE", "04": "Transferencia / SINPE",
              "05": "Transferencia / SINPE", "06": "Transferencia / SINPE", "07": "Transferencia / SINPE"}


def norm(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip()


def slug(s, n=40):
    return re.sub(r"[^A-Za-z0-9]+", "-", norm(s)).strip("-")[:n] or "SinNombre"


def local(tag):
    return tag.rsplit("}", 1)[-1]


def find(el, path):
    """Busca por nombre local, ignorando namespaces. path = 'A/B/C'."""
    cur = [el]
    for part in path.split("/"):
        nxt = []
        for c in cur:
            nxt += [x for x in c.iter() if local(x.tag) == part and x is not c]
        if not nxt:
            return None
        cur = nxt[:1]
    return cur[0]


def txt(el, path, default=""):
    x = find(el, path)
    return (x.text or "").strip() if x is not None and x.text else default


def parse_xml(p):
    try:
        root = ET.parse(p).getroot()
    except ET.ParseError:
        return None
    tipo = local(root.tag)
    if tipo not in ("FacturaElectronica", "TiqueteElectronico", "NotaCreditoElectronica",
                    "NotaDebitoElectronica", "FacturaElectronicaCompra"):
        return None
    emisor = find(root, "Emisor")
    lineas = [txt(l, "Detalle") for l in root.iter() if local(l.tag) == "LineaDetalle"]
    medio = find(root, "TipoMedioPago") or find(root, "MedioPago")
    total = float(txt(root, "ResumenFactura/TotalComprobante", "0") or 0)
    if tipo == "NotaCreditoElectronica":
        total = -total
    return {
        "clave": txt(root, "Clave"),
        "consecutivo": txt(root, "NumeroConsecutivo"),
        "tipo_doc": tipo,
        "fecha": txt(root, "FechaEmision")[:10],
        "proveedor": txt(emisor, "Nombre") if emisor is not None else "",
        "nombre_comercial": txt(emisor, "NombreComercial") if emisor is not None else "",
        "cedula_emisor": txt(emisor, "Identificacion/Numero") if emisor is not None else "",
        "detalle": lineas,
        "moneda": txt(root, "CodigoMoneda", "CRC") or "CRC",
        "tipo_cambio_doc": float(txt(root, "TipoCambio", "1") or 1),
        "total": total,
        "medio_pago": (medio.text or "").strip() if medio is not None and medio.text else "",
    }


NUM = r"(?:₡|¢|CRC|\$|USD)?\s*([0-9]{1,3}(?:[.,\s][0-9]{3})*(?:[.,][0-9]{1,2})|[0-9]+(?:[.,][0-9]{1,2})?)"


def to_float(s):
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[-1]) in (1, 2) else s.replace(",", "")
    elif s.count(".") > 1 or (s.count(".") == 1 and len(s.split(".")[-1]) == 3):
        s = s.replace(".", "")
    return float(s)


def parse_pdf(p):
    """Respaldo cuando no viene XML: extrae fecha, total y emisor del texto del PDF."""
    try:
        with pdfplumber.open(p) as pdf:
            text = "\n".join((pg.extract_text() or "") for pg in pdf.pages[:3])
    except Exception:
        return None
    out = {"texto": text}
    m = re.search(r"\b(506\d{47})\b", text)
    out["clave"] = m.group(1) if m else ""
    m = re.search(r"\b(\d{20})\b", text)
    out["consecutivo"] = m.group(1) if m else ""
    totals = re.findall(r"(?i)total\s*(?:comprobante|factura|a pagar|general|documento)?\s*[:]?\s*" + NUM, text)
    vals = []
    for t in totals:
        try:
            vals.append(to_float(t))
        except ValueError:
            pass
    out["total"] = max(vals) if vals else None
    out["moneda"] = "USD" if re.search(r"(?i)\bUSD\b|d[oó]lar|US\$", text) and not re.search(r"₡|CRC|colon", text) else "CRC"
    fechas = re.findall(r"\b(\d{1,2})[/-](\d{1,2})[/-](20\d{2})\b", text) or []
    iso = re.findall(r"\b(20\d{2})-(\d{2})-(\d{2})\b", text)
    f = None
    for d, mth, y in fechas:
        try:
            f = date(int(y), int(mth), int(d)); break
        except ValueError:
            continue
    if f is None and iso:
        y, mth, d = iso[0]
        try:
            f = date(int(y), int(mth), int(d))
        except ValueError:
            pass
    out["fecha"] = f.isoformat() if f else ""
    return out


def categoria_y_desc(prov, comercial, detalle, hist):
    key = norm(f"{prov} {comercial}").lower()
    for hp, (cat, desc) in hist.items():
        if hp and (hp in key or key[:12] in hp):
            return cat, desc, False
    blob = norm(f"{prov} {comercial} {' '.join(detalle[:5])}").lower()
    for pat, cat in REGLAS:
        if re.search(pat, blob):
            return cat, (detalle[0][:60] if detalle else cat), False
    return "Imprevistos y otros", (detalle[0][:60] if detalle else "Factura"), True


def historial(ws):
    """Proveedor -> (categoría, descripción) aprendido de los registros manuales existentes."""
    h = {}
    for r in range(6, ws.max_row + 1):
        prov, cat, desc = ws.cell(r, 4).value, ws.cell(r, 2).value, ws.cell(r, 3).value
        if prov and cat and len(norm(str(prov))) >= 4 and str(prov) not in ("Sin Factura", "Canavida"):
            h[norm(str(prov)).lower()] = (cat, desc)
    return h


def existentes(ws):
    s = set()
    for r in range(6, ws.max_row + 1):
        g = ws.cell(r, 7).value
        if isinstance(g, (int, float)):
            s.add(int(g))
    return s


def load_ledger():
    if not LEDGER.exists():
        return []
    with LEDGER.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def save_ledger(rows):
    LEDGER.parent.mkdir(parents=True, exist_ok=True)
    rows.sort(key=lambda r: (r["fecha"], r["proveedor"]))
    with LEDGER.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)


def procesar_inbox(ledger, hist, ya_en_hoja, hoy):
    vistos_msg = {r["message_id"] for r in ledger}
    vistas_claves = {r["clave"] for r in ledger if r["clave"]}
    nuevas = []
    for d in sorted(p for p in INBOX.iterdir() if p.is_dir()):
        mid = d.name
        if mid in vistos_msg:
            continue
        meta = json.loads((d / "meta.json").read_text(encoding="utf-8")) if (d / "meta.json").exists() else {}
        pdfs = sorted(d.glob("*.pdf")) + sorted(d.glob("*.PDF"))
        xmls = [x for x in sorted(d.glob("*.xml")) + sorted(d.glob("*.XML")) if "respuesta" not in x.name.lower()]
        docs = [x for x in (parse_xml(x) for x in xmls) if x]
        if not pdfs:
            ledger.append(dict.fromkeys(FIELDS, "") | {"message_id": mid, "estado": "sin_pdf",
                          "notas": (meta.get("subject") or "")[:120], "procesado": hoy})
            continue
        if not docs:  # sin XML: leer el PDF
            for p in pdfs:
                info = parse_pdf(p) or {}
                subj = meta.get("subject", "")
                docs.append({"clave": info.get("clave", ""), "consecutivo": info.get("consecutivo", ""),
                             "tipo_doc": "PDF", "fecha": info.get("fecha") or (meta.get("date") or "")[:10],
                             "proveedor": re.sub(r"^(RV|RE|FW|Fwd):\s*", "", subj, flags=re.I)[:60] or p.stem,
                             "nombre_comercial": "", "cedula_emisor": "", "detalle": [],
                             "moneda": info.get("moneda", "CRC"), "tipo_cambio_doc": 1.0,
                             "total": info.get("total"), "medio_pago": "", "_pdf": p, "_texto": info.get("texto", "")})
        for i, doc in enumerate(docs):
            pdf = doc.pop("_pdf", None) or (next((p for p in pdfs if doc["clave"] and doc["clave"] in p.name), None)
                                            or pdfs[min(i, len(pdfs) - 1)])
            row = dict.fromkeys(FIELDS, "")
            row.update(message_id=mid if i == 0 else f"{mid}#{i}", procesado=hoy,
                       clave=doc["clave"], consecutivo=doc["consecutivo"], tipo_doc=doc["tipo_doc"],
                       fecha=doc["fecha"], proveedor=doc["proveedor"], nombre_comercial=doc["nombre_comercial"],
                       cedula_emisor=doc["cedula_emisor"], moneda=doc["moneda"])
            notas = []
            if doc["cedula_emisor"] == CANAVIDA_ID:
                row.update(estado="emitida_por_canavida", notas="Factura emitida por Canavida (ingreso, no gasto)")
            elif doc["clave"] and doc["clave"] in vistas_claves:
                row.update(estado="duplicada", notas="Misma clave que otra factura ya registrada")
            else:
                row["estado"] = "registrada"
            if doc["clave"]:
                vistas_claves.add(doc["clave"])
            total = doc["total"]
            if total is None:
                row["estado"] = "revisar"
                notas.append("No se pudo leer el monto del PDF")
                total = 0.0
            row["total_moneda"] = f"{total:.2f}"
            if doc["moneda"] == "USD":
                row["tipo_cambio"], usd = "1", total
            else:
                tc = doc["tipo_cambio_doc"] if doc["moneda"] not in ("CRC", "") else TIPO_CAMBIO
                tc = TIPO_CAMBIO if doc["moneda"] in ("CRC", "") else tc
                row["tipo_cambio"], usd = f"{tc:g}", total / tc
            row["monto_usd"] = f"{usd:.2f}"
            cat, desc, dudosa = categoria_y_desc(doc["proveedor"], doc["nombre_comercial"], doc["detalle"], hist)
            row["categoria"], row["descripcion"] = cat, desc
            if dudosa:
                notas.append("Revisar categoría")
            row["metodo_pago"] = MEDIO_PAGO.get(doc["medio_pago"], "Tarjeta")
            try:
                f = date.fromisoformat(row["fecha"])
            except ValueError:
                f = None
                notas.append("Fecha no legible")
                row["estado"] = "revisar"
            if f and not (FECHA_MIN <= f <= FECHA_MAX):
                notas.append("Fecha fuera del periodo de la hoja")
                if row["estado"] == "registrada":
                    row["estado"] = "fuera_periodo"
            cons_int = int(doc["consecutivo"][-10:]) if doc["consecutivo"][-10:].isdigit() else None
            if row["estado"] == "registrada" and cons_int and cons_int in ya_en_hoja and f and f <= date(2026, 10, 31):
                row["estado"] = "ya_en_hoja"
                notas.append("Nº de factura ya aparece en la hoja")
            row["mes"] = f.strftime("%Y-%m") if f else "sin-fecha"
            # archivar PDF por mes
            dest_dir = FACT_DIR / row["mes"]
            dest_dir.mkdir(parents=True, exist_ok=True)
            nombre = f"{row['fecha'] or 'sin-fecha'}_{slug(doc['nombre_comercial'] or doc['proveedor'])}_{(doc['consecutivo'] or pdf.stem)[-10:]}.pdf"
            if row["estado"] != "duplicada":
                shutil.copy2(pdf, dest_dir / nombre)
                row["pdf"] = f"facturas/{row['mes']}/{nombre}"
            row["fuente"] = "xml" if doc["tipo_doc"] != "PDF" else "pdf"
            row["notas"] = "; ".join(filter(None, [row["notas"]] + notas))
            ledger.append(row)
            nuevas.append(row)
    return nuevas


def escribir_excel(ledger, out):
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Gastos Diarios"]
    # primera fila libre (B..F vacías) después de los registros manuales
    r = 6
    last = 5
    for i in range(6, ws.max_row + 1):
        if any(ws.cell(i, c).value not in (None, "") for c in range(2, 7)):
            last = i
    r = last + 1
    estilo = [copy(ws.cell(last, c)._style) for c in range(1, 9)]
    fmt = [ws.cell(last, c).number_format for c in range(1, 9)]
    for row in sorted((x for x in ledger if x["estado"] == "registrada"), key=lambda x: x["fecha"]):
        vals = [datetime.fromisoformat(row["fecha"]), row["categoria"], row["descripcion"],
                (row["nombre_comercial"] or row["proveedor"])[:60], row["metodo_pago"],
                (f"={float(row['total_moneda']):.2f}/{float(row['tipo_cambio']):g}" if row["moneda"] in ("CRC", "")
                 else float(row["monto_usd"])),
                int(row["consecutivo"][-10:]) if row["consecutivo"][-10:].isdigit() else row["consecutivo"],
                "Auto Gmail" + (f" · {row['notas']}" if row["notas"] else "")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell._style = copy(estilo[c - 1])
            cell.number_format = fmt[c - 1]
        r += 1
    wb.save(out)


def leer_hoja_manual():
    """Gastos manuales existentes (para que el resumen PDF refleje toda la hoja)."""
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Gastos Diarios"]
    out = []
    for i in range(6, ws.max_row + 1):
        a, b, f = ws.cell(i, 1).value, ws.cell(i, 2).value, ws.cell(i, 6).value
        if not (isinstance(a, datetime) and b and f not in (None, "")):
            continue
        if isinstance(f, str) and f.startswith("="):
            expr = f[1:]
            if not re.fullmatch(r"[0-9.+\-*/() ]+", expr):
                continue
            f = eval(expr, {"__builtins__": {}})
        out.append({"fecha": a.date().isoformat(), "categoria": b, "monto_usd": float(f)})
    return out


def escribir_pdf(ledger, nuevas, manual, out, hoy):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

    ss = getSampleStyleSheet()
    small = ss["BodyText"].clone("s", fontSize=7.5, leading=9)
    doc = SimpleDocTemplate(str(out), pagesize=landscape(letter), leftMargin=1.2 * cm, rightMargin=1.2 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm, title="Reporte de facturas Canavida")
    verde = colors.HexColor("#1f5f4a")
    ts = TableStyle([("BACKGROUND", (0, 0), (-1, 0), verde), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                     ("FONTSIZE", (0, 0), (-1, -1), 7.5), ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                     ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f3")]),
                     ("VALIGN", (0, 0), (-1, -1), "TOP")])
    reg = [x for x in ledger if x["estado"] == "registrada"]
    el = [Paragraph("Canavida Villas &amp; Resort — Reporte diario de facturas", ss["Title"]),
          Paragraph(f"Generado el {hoy} · Facturas PDF extraídas del correo · Montos en USD (₡ a {TIPO_CAMBIO:g})", ss["Normal"]),
          Spacer(1, 8)]

    # Resumen por mes (manual + automático)
    meses = sorted({x["fecha"][:7] for x in reg} | {x["fecha"][:7] for x in manual})
    cats = sorted({x["categoria"] for x in reg} | {x["categoria"] for x in manual})
    tot = defaultdict(float)
    for x in manual:
        tot[(x["categoria"], x["fecha"][:7], "m")] += x["monto_usd"]
    for x in reg:
        tot[(x["categoria"], x["fecha"][:7], "a")] += float(x["monto_usd"])
    nombre_mes = lambda m: f"{MESES[int(m[5:])]} {m[:4]}"
    data = [["Categoría"] + [nombre_mes(m) for m in meses] + ["Total"]]
    for c in cats:
        fila = [tot[(c, m, "m")] + tot[(c, m, "a")] for m in meses]
        data.append([c] + [f"${v:,.2f}" for v in fila] + [f"${sum(fila):,.2f}"])
    col = [sum(tot[(c, m, k)] for c in cats for k in "ma") for m in meses]
    data.append(["TOTAL GASTOS"] + [f"${v:,.2f}" for v in col] + [f"${sum(col):,.2f}"])
    t = Table(data, repeatRows=1)
    t.setStyle(ts)
    t.setStyle(TableStyle([("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
    el += [Paragraph("Resumen de gastos por mes y categoría (registros manuales + facturas del correo)", ss["Heading2"]), t,
           Spacer(1, 10)]

    # Nuevas hoy
    el.append(Paragraph(f"Facturas nuevas en esta actualización: {len([n for n in nuevas if n['estado']=='registrada'])}", ss["Heading2"]))
    if nuevas:
        d = [["Fecha", "Proveedor", "Nº", "Categoría", "Moneda", "Total", "USD", "Estado / notas"]]
        for n in sorted(nuevas, key=lambda x: x["fecha"]):
            d.append([n["fecha"], Paragraph((n["nombre_comercial"] or n["proveedor"])[:45], small), n["consecutivo"][-10:],
                      n["categoria"], n["moneda"], f"{float(n['total_moneda'] or 0):,.2f}", f"${float(n['monto_usd'] or 0):,.2f}",
                      Paragraph(f"{n['estado']} {n['notas']}"[:120], small)])
        t = Table(d, repeatRows=1, colWidths=[2 * cm, 6 * cm, 2.2 * cm, 4.2 * cm, 1.4 * cm, 2.4 * cm, 2.2 * cm, 5.5 * cm])
        t.setStyle(ts)
        el.append(t)
    else:
        el.append(Paragraph("No llegaron facturas nuevas desde la última actualización.", ss["Normal"]))

    # Detalle por mes
    for m in sorted({x["mes"] for x in reg}):
        filas = sorted((x for x in reg if x["mes"] == m), key=lambda x: x["fecha"])
        el += [PageBreak(), Paragraph(f"{nombre_mes(m)} — {len(filas)} facturas del correo · "
                                       f"${sum(float(x['monto_usd']) for x in filas):,.2f}", ss["Heading2"])]
        d = [["Fecha", "Proveedor", "Descripción", "Categoría", "Pago", "Total", "USD", "Archivo"]]
        for x in filas:
            d.append([x["fecha"], Paragraph((x["nombre_comercial"] or x["proveedor"])[:45], small),
                      Paragraph(x["descripcion"][:50], small), Paragraph(x["categoria"], small), x["metodo_pago"][:13],
                      f"{x['moneda']} {float(x['total_moneda']):,.2f}", f"${float(x['monto_usd']):,.2f}",
                      Paragraph(x["pdf"].split("/")[-1][:48], small)])
        t = Table(d, repeatRows=1, colWidths=[1.9 * cm, 4.8 * cm, 4.2 * cm, 3.6 * cm, 2.6 * cm, 3 * cm, 2 * cm, 4.8 * cm])
        t.setStyle(ts)
        el.append(t)

    otros = [x for x in ledger if x["estado"] not in ("registrada", "sin_pdf")]
    if otros:
        el += [PageBreak(), Paragraph("Documentos no sumados como gasto (revisar)", ss["Heading2"])]
        d = [["Fecha", "Proveedor", "Monto", "Estado", "Notas"]]
        for x in sorted(otros, key=lambda x: x["fecha"]):
            d.append([x["fecha"], Paragraph((x["nombre_comercial"] or x["proveedor"])[:50], small),
                      f"{x['moneda']} {float(x['total_moneda'] or 0):,.2f}", x["estado"], Paragraph(x["notas"][:120], small)])
        t = Table(d, repeatRows=1, colWidths=[2 * cm, 7 * cm, 3.5 * cm, 3.5 * cm, 9 * cm])
        t.setStyle(ts)
        el.append(t)
    doc.build(el)


def main():
    hoy = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Gastos Diarios"]
    ledger = load_ledger()
    nuevas = procesar_inbox(ledger, historial(ws), existentes(ws), hoy) if INBOX.exists() else []
    save_ledger(ledger)
    REP_DIR.mkdir(exist_ok=True)
    xlsx = REP_DIR / f"Canavida_Gastos_Ingresos_{hoy}.xlsx"
    pdf = REP_DIR / f"Reporte_Facturas_{hoy}.pdf"
    escribir_excel(ledger, xlsx)
    escribir_pdf(ledger, nuevas, leer_hoja_manual(), pdf, hoy)
    reg = [x for x in nuevas if x["estado"] == "registrada"]
    resumen = {"nuevas_registradas": len(reg), "nuevas_total_usd": round(sum(float(x["monto_usd"]) for x in reg), 2),
               "otras_nuevas": {k: sum(1 for x in nuevas if x["estado"] == k) for k in {x["estado"] for x in nuevas} - {"registrada"}},
               "por_mes": {m: round(sum(float(x["monto_usd"]) for x in ledger if x["estado"] == "registrada" and x["mes"] == m), 2)
                           for m in sorted({x["mes"] for x in ledger if x["estado"] == "registrada"})},
               "revisar_categoria": sum(1 for x in reg if "Revisar categoría" in x["notas"]),
               "xlsx": str(xlsx), "pdf": str(pdf)}
    print(json.dumps(resumen, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
