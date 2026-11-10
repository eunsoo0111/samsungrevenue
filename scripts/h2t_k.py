import sys, re
from bs4 import BeautifulSoup
p = sys.argv[1]
raw = open(p,'rb').read()
for enc in ('utf-8','euc-kr','cp949'):
    try: txt = raw.decode(enc); break
    except: pass
s = BeautifulSoup(txt, 'lxml')
for t in s(['script','style']): t.decompose()
def cell(c): return re.sub(r'\s+',' ', c.get_text(' ', strip=True))
for el in (s.body or s).find_all(['p','table'], recursive=True):
    if el.name=='table':
        if el.find_parent('table'): continue
        for tr in el.find_all('tr'):
            print("  | " + " | ".join(cell(c) for c in tr.find_all(['td','th'])))
    elif not el.find_parent('table'):
        t=cell(el)
        if t: print(t)
