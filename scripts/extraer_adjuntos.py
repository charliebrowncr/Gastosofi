import json,glob,base64,email,os,re,sys
from email import policy
D="/root/.claude/projects/-home-user-Gastosofi/5f64594f-d4e4-558c-aa52-fe0460172950/tool-results/"
files=sorted(glob.glob(D+"mcp-Google-Drive-download_file_content-17934*.txt"))
idx=[]
def walk(m):
    for p in m.walk():
        if p.get_content_type()=="message/rfc822":
            continue
        fn=p.get_filename()
        if fn and not p.is_multipart():
            d=p.get_payload(decode=True)
            if d: yield fn,d
for f in files:
    j=json.load(open(f)); t=j["title"]; fid=j["id"]
    raw=base64.b64decode(j["content"]+"==")
    m=email.message_from_bytes(raw,policy=policy.default)
    out=f"em/{fid}"; os.makedirs(out,exist_ok=True)
    got=[]
    # nested rfc822 too
    msgs=[m]+[p.get_payload()[0] for p in m.walk() if p.get_content_type()=="message/rfc822" and p.get_payload()]
    for mm in msgs:
        for fn,d in walk(mm):
            fn=os.path.basename(fn); open(os.path.join(out,fn),"wb").write(d); got.append(fn)
    idx.append(dict(id=fid,title=t,subject=str(m["subject"]),frm=str(m["from"]),date=str(m["date"]),files=sorted(set(got))))
json.dump(idx,open("em_index.json","w"),ensure_ascii=False,indent=1)
for i in idx: print(i["title"][:70],"|",i["frm"][:30],"|",[x for x in i["files"] if x.lower().endswith((".pdf",".xml"))][:6])
