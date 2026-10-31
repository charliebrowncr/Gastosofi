import json,os,shutil,re
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill
from reportlab.lib.pagesizes import A4,landscape
from reportlab.platypus import SimpleDocTemplate,Table,TableStyle,Paragraph,Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
R=json.load(open("rows.json"))
IDX=json.load(open("em_index.json"))
for _i,_r in enumerate(R): _r["eid"]=IDX[_i]["id"]
L=[]  # dict rows
seen=set()
def add(**k):
    d=dict(tipo="Factura",sub=None,iva=None,desc="",pdf="",mon="CRC",cuenta=True); d.update(k); L.append(d)
# previous 3 (factoc)
add(fecha="2026-10-14",num="00100090010000090990",prov="MOSAIC WINE PARADISE, SRL",sub=45878.05,iva=5964.15,total=56430.00,pdf="50614102600310274744200100090010000090990109512219.pdf",src="in")
add(fecha="2026-10-22",num="00100001010000020687",prov="LAURA MARIA RIVERA FALLAS",sub=6902.65,iva=897.35,total=7800.00,pdf="000110690289-FE-00100001010000020687.pdf",src="in")
add(fecha="2026-10-26",num="00100001010000020713",prov="LAURA MARIA RIVERA FALLAS",sub=12252.26,iva=1497.74,total=13750.00,pdf="000110690289-FE-00100001010000020713.pdf",src="in")
for r in R:
    t=r["src"]; p=r["pdfs"][0] if r["pdfs"] else ""
    base=dict(pdf=p,src=None,id=None)
    if r["total"] is not None:
        key=r["num"]+r["prov"]
        if key in seen: continue
        seen.add(key)
        add(fecha=r["fecha"],num=r["num"],prov=r["prov"],mon=r["mon"],sub=r["sub"],iva=r["iva"],total=r["total"],pdf=p,src="em",id=r["src"])
        L[-1]["dir"]=r
    elif "0087" in t:
        add(fecha="2026-10-30",num="00100001010000000087",prov="MARIO ARIAS SEGURA",sub=2292699.12,iva=298050.88,total=2590750.00,pdf=p,src="em",desc="Venta e instalación de puertas y ventanas"); L[-1]["dir"]=r
    elif "ICE" in t:
        add(tipo="Contrato",fecha="2026-10-26",num="21656013",prov="Instituto Costarricense de Electricidad (ICE)",total=None,pdf="Cont_Serv_Postpago_3102777726_21656013_26-10-2026_08-36-48-AM.pdf",src="em",desc="Contratos del ICE",cuenta=False); L[-1]["dir"]=r
    elif "159109" in t:
        add(tipo="Pre-factura",fecha="2026-09-04",num="44004",prov="CRC ACCOUNTING GROUP SRL",mon="USD",sub=297.35,iva=38.65,total=336.00,pdf=p,src="em",desc="Informe de aseguramiento origen de fondos - Banco Nacional",cuenta=False); L[-1]["dir"]=r
    elif p=="RC-2462.pdf":
        add(tipo="Recibo",fecha="2026-10-27",num="RC-2462",prov="CLEAN VITA CR S.R.L. (Natural Cleaner)",total=3955.00,pdf=p,src="em",desc="Pago de factura 3064 (no se suma, ya incluida)",cuenta=False); L[-1]["dir"]=r
    elif p=="RC-2461.pdf":
        add(tipo="Recibo",fecha="2026-10-27",num="RC-2461",prov="CLEAN VITA CR S.R.L. (Natural Cleaner)",total=248600.00,pdf=p,src="em",desc="Pago de factura 3059 (no se suma, ya incluida)",cuenta=False); L[-1]["dir"]=r
    elif p=="Comprobante.pdf":
        add(tipo="Comprobante de pago",fecha="2026-10-26",num="R202610263467652",prov="INSTITUTO NACIONAL DE SEGUROS",mon="USD",total=245.34,pdf=p,src="em",desc="Cargo automático póliza 0220GFL001306802 (no se suma, ya incluido en su factura)",cuenta=False); L[-1]["dir"]=r
    elif "Bit-Office" in t:
        add(tipo="Recibo",fecha="2026-10-26",num="73207952-171458683",prov="Bit-Office (Paddle)",mon="USD",total=25.50,pdf="",src=None,desc="Recibo sin PDF (solo texto del correo); no se suma",cuenta=False)
# sort
L.sort(key=lambda x:(x["fecha"],x["num"]))
# folders with PDFs by month
shutil.rmtree("facturas",ignore_errors=True)
def safe(s): return re.sub(r"[^A-Za-z0-9._-]+","_",s)[:40]
for x in L:
    x["mes"]=x["fecha"][:7]
    if not x["pdf"] or not x.get("src"): continue
    d=f"facturas/{x['mes']}"; os.makedirs(d,exist_ok=True)
    srcdir="in" if x["src"]=="in" else "em/"+x["dir"]["eid"]
    sp=os.path.join(srcdir,x["pdf"])
    if os.path.exists(sp):
        nm=f"{x['fecha']}_{safe(x['prov'])}_{safe(x['num'])}.pdf"
        shutil.copy(sp,os.path.join(d,nm)); x["file"]=nm
h=["Mes","Fecha","Tipo","N°","Proveedor","Moneda","Subtotal","IVA","Total","Descripción","PDF"]
wb=Workbook();ws=wb.active;ws.title="Facturas"
ws.append(h)
for c in ws[1]: c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor="DDDDDD")
for x in L: ws.append([x["mes"],x["fecha"],x["tipo"],x["num"],x["prov"],x["mon"],x["sub"],x["iva"],x["total"],x["desc"],x.get("file","")])
for col,w in zip("ABCDEFGHIJK",[9,11,18,24,40,8,14,12,14,52,50]): ws.column_dimensions[col].width=w
for row in ws.iter_rows(min_row=2):
    for c in row[6:9]: c.number_format='#,##0.00'
ws2=wb.create_sheet("Resumen por mes")
ws2.append(["Mes","Moneda","Facturas","Total"])
for c in ws2[1]: c.font=Font(bold=True)
agg={}
for x in L:
    if x["tipo"]=="Factura" and x["total"] is not None:
        k=(x["mes"],x["mon"]); a=agg.setdefault(k,[0,0]); a[0]+=1; a[1]+=x["total"]
for (m,mo),(n,t) in sorted(agg.items()): ws2.append([m,mo,n,round(t,2)])
for col,w in zip("ABCD",[10,8,10,16]): ws2.column_dimensions[col].width=w
for row in ws2.iter_rows(min_row=2): row[3].number_format='#,##0.00'
wb.save("facturas_octubre_2026.xlsx")
st=getSampleStyleSheet(); sm=st["Normal"].clone("sm"); sm.fontSize=6.5; sm.leading=8
fm=lambda v:"" if v is None else f"{v:,.2f}"
data=[["Fecha","Tipo","N°","Proveedor","Mon.","Total","Descripción"]]+[[x["fecha"],x["tipo"],x["num"][-12:],Paragraph(x["prov"],sm),x["mon"],fm(x["total"]),Paragraph(x["desc"],sm)] for x in L]
t=Table(data,repeatRows=1,colWidths=[55,60,75,170,28,70,300]);t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),.3,colors.grey),("FONTSIZE",(0,0),(-1,-1),6.5),("VALIGN",(0,0),(-1,-1),"TOP"),("ALIGN",(5,1),(5,-1),"RIGHT")]))
tot=[["Mes","Moneda","Facturas","Total"]]+[[m,mo,n,fm(t_)] for (m,mo),(n,t_) in sorted(agg.items())]
t2=Table(tot,colWidths=[60,50,50,90]);t2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("GRID",(0,0),(-1,-1),.3,colors.grey),("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(3,1),(3,-1),"RIGHT")]))
SimpleDocTemplate("resumen_facturas_octubre_2026.pdf",pagesize=landscape(A4),leftMargin=24,rightMargin=24,topMargin=24,bottomMargin=24).build([Paragraph("Canavida - Facturas octubre 2026 (actualización 31/10/2026)",st["Title"]),Paragraph("Totales por mes (solo facturas; recibos, comprobantes, contrato y pre-factura se listan pero no se suman):",st["Normal"]),Spacer(1,6),t2,Spacer(1,10),t])
print(len(L),agg)
for x in L: print(x["fecha"],x["tipo"][:8],x["num"][-10:],x["prov"][:28],x["mon"],x["total"],x.get("file","-")[:40])
