import re, sys, os, time, subprocess, urllib.parse
OUT=sys.argv[1]
def get(url, path):
    if os.path.exists(path) and os.path.getsize(path)>500: return open(path,encoding='utf-8',errors='ignore').read()
    for i in range(5):
        r=subprocess.run(["curl","-sS","--max-time","60","-A","Mozilla/5.0",url,"-o",path],capture_output=True)
        if r.returncode==0 and os.path.exists(path) and os.path.getsize(path)>200: break
        time.sleep(2**i)
    time.sleep(1)
    return open(path,encoding='utf-8',errors='ignore').read()
for rcp in sys.argv[2:]:
    h=get(f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcp}", f"{OUT}/{rcp}_main.html")
    m=re.search(r'viewDoc\("(\d+)", "(\d+)", "(\d*)", "(\d*)", "(\d*)", "([^"]*)"', h)
    if not m:
        m2=re.search(r"node\d\['dcmNo'\] = \"(\d+)\";\s*node\d\['eleId'\] = \"(\d+)\";\s*node\d\['offset'\] = \"(\d+)\";\s*node\d\['length'\] = \"(\d+)\";\s*node\d\['dtd'\] = \"(.*?)\";",h)
        if not m2: print(rcp,'no doc'); continue
        args=(rcp,)+m2.groups()
    else: args=m.groups()
    q=urllib.parse.urlencode(dict(rcpNo=args[0],dcmNo=args[1],eleId=args[2],offset=args[3],length=args[4],dtd=args[5]))
    get(f"https://dart.fss.or.kr/report/viewer.do?{q}", f"{OUT}/{rcp}.html")
    print(rcp,'ok',flush=True)
