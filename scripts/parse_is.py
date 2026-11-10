import re, glob, os, json, sys
BASE=sys.argv[1]
ITEMS=[('매출액','매출액'),('매출원가','매출원가'),('매출총이익','매출총이익'),('판매비와관리비','판매비와관리비'),('영업이익','영업이익'),
       ('금융수익','금융수익'),('금융원가','금융원가'),('지분법','지분법손익'),('기타수익','기타수익'),('기타비용','기타비용'),
       ('법인세비용차감전','법인세비용차감전순이익'),('법인세비용 ','법인세비용'),('계속영업','계속영업당기순이익'),('중단영업','중단영업손익'),('당기순이익','당기순이익')]
def num(s):
    s=s.strip().replace(',','')
    neg=s.startswith('(')
    s=s.strip('()')
    try: v=float(s)
    except: return None
    return -v if neg else v
out={}
for d in sorted(glob.glob(os.path.join(BASE,'2*/'))):
    rcp=os.path.basename(d.rstrip('/'))
    fs=[f for f in glob.glob(d+'*연결재무제표.txt') if '주석' not in f]
    if not fs: continue
    t=open(fs[0]).read()
    i=t.find('포괄손익계산서'); 
    if i<0: i=t.find('손익계산서')
    seg=t[i:i+12000]
    res={}; ctrl=False; got_parent=False
    for line in seg.split('\n'):
        if not line.startswith('  | '): continue
        c=[x.strip() for x in line[4:].split(' | ')]
        lab=c[0]
        vals=[num(x) for x in c[1:]]
        vals=[v for v in vals if v is not None]
        if not vals: 
            if '당기순이익의 귀속' in lab: ctrl=True
            continue
        key=None
        for k,name in ITEMS:
            if lab.startswith(k) or (k=='법인세비용 ' and lab.startswith('법인세비용') and '차감전' not in lab):
                key=name; break
        if lab.startswith('수익(매출액)'): key='매출액'
        if key and key not in res: res[key]=vals
        if ctrl and lab.startswith('지배기업') and not got_parent:
            res['지배주주순이익']=vals; got_parent=True
    out[rcp]=res
json.dump(out,open(os.path.join(BASE,'is.json'),'w'),ensure_ascii=False,indent=0)
for k in ['20240329002895','20250311001190','20260310003071','20250515002323','20250814004371','20251114002848','20260515002842','20260814003805','20230307000653','20220308000667']:
    print(k, {a:[round(x/1e8) for x in b] for a,b in out.get(k,{}).items()})
