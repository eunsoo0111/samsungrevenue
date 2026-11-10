import re, sys, os, json, time, subprocess, urllib.parse
OUT = sys.argv[1]
rcps = sys.argv[2:]
UA = "Mozilla/5.0"
def get(url, path):
    if os.path.exists(path) and os.path.getsize(path) > 500: return open(path, encoding='utf-8', errors='ignore').read()
    for i in range(4):
        r = subprocess.run(["curl","-sS","--max-time","60","-A",UA,url,"-o",path])
        if r.returncode==0 and os.path.getsize(path)>200: break
        time.sleep(2**i)
    time.sleep(0.7)
    return open(path, encoding='utf-8', errors='ignore').read()
WANT = ["사업의 내용","주요 제품","원재료","매출 및 수주","연결재무제표","요약재무","사업의 개요","부문"]
for rcp in rcps:
    d = os.path.join(OUT, rcp); os.makedirs(d, exist_ok=True)
    html = get(f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp}", os.path.join(d,"main.html"))
    nodes = []
    for m in re.finditer(r"node\d\['text'\] = \"(.*?)\";.*?node\d\['dcmNo'\] = \"(\d+)\";\s*node\d\['eleId'\] = \"(\d+)\";\s*node\d\['offset'\] = \"(\d+)\";\s*node\d\['length'\] = \"(\d+)\";\s*node\d\['dtd'\] = \"(.*?)\";", html, re.S):
        nodes.append(dict(text=m.group(1).strip(), dcmNo=m.group(2), eleId=m.group(3), offset=m.group(4), length=m.group(5), dtd=m.group(6)))
    json.dump(nodes, open(os.path.join(d,"toc.json"),"w"), ensure_ascii=False, indent=1)
    for n in nodes:
        if any(w in n['text'] for w in WANT) and '별도' not in n['text']:
            q = urllib.parse.urlencode(dict(rcpNo=rcp, dcmNo=n['dcmNo'], eleId=n['eleId'], offset=n['offset'], length=n['length'], dtd=n['dtd']))
            safe = re.sub(r'[^\w가-힣]+','_', n['text'])[:40]
            get(f"https://dart.fss.or.kr/report/viewer.do?{q}", os.path.join(d, f"{int(n['eleId']):03d}_{safe}.html"))
    print(rcp, len(nodes), "nodes")
