import json,glob,os,re,subprocess
import xml.etree.ElementTree as ET
idx=json.load(open("em_index.json"))
rows=[]
def first(r,k):
    for e in r.iter():
        if e.tag.split('}')[-1]==k and (e.text or "").strip(): return e.text.strip()
for i in idx:
    d=f"em/{i['id']}"
    pdfs=[f for f in i["files"] if f.lower().endswith(".pdf")]
    xmls=[f for f in i["files"] if f.lower().endswith(".xml") and not re.search(r"respuesta|^MH|^AHC|^Res-|HAFA|_F\.xml|^HA",f,re.I)]
    info=dict(src=i["title"],pdfs=pdfs,prov="",fecha="",num="",mon="",total=None,sub=None,iva=None,desc="")
    for x in xmls:
        try:
            r=ET.parse(os.path.join(d,x)).getroot()
        except Exception: continue
        if not first(r,"TotalComprobante"): continue
        info["num"]=first(r,"NumeroConsecutivo") or ""
        em=[e for e in r if e.tag.split('}')[-1]=="Emisor"]
        info["prov"]=(first(em[0],"Nombre") if em else first(r,"Nombre")) or ""
        info["fecha"]=(first(r,"FechaEmision") or "")[:10]
        info["mon"]=first(r,"CodigoMoneda") or "CRC"
        info["total"]=round(float(first(r,"TotalComprobante")),2)
        info["iva"]=round(float(first(r,"TotalImpuesto") or 0),2)
        info["sub"]=round(float(first(r,"TotalVenta") or 0)-float(first(r,"TotalDescuentos") or 0),2)
        break
    txt=""
    for p in pdfs:
        t=subprocess.run(["pdftotext","-layout",os.path.join(d,p),"-"],capture_output=True,text=True).stdout
        txt+=f"\n##{p}\n"+t
    info["txt"]=txt
    rows.append(info)
json.dump(rows,open("rows.json","w"),ensure_ascii=False,indent=1)
for r in rows: print(r["fecha"],r["num"][-8:],r["prov"][:34],r["mon"],r["total"],"|",(r["pdfs"] or ["-"])[0][:30])
