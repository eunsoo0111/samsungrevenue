import re, os, json, glob, sys
BASE = sys.argv[1]
SEGMAP = {'컴포넌트':'COMP','광학통신솔루션':'OPT','광학솔루션':'OPT','모듈':'OPT','패키지솔루션':'PKG','기판':'PKG'}
def num(s):
    s=s.strip().replace(',','')
    neg = s.startswith('(') and s.endswith(')')
    s=s.strip('()')
    if s in ('','-'): return None
    try: v=float(s)
    except: return None
    return -v if neg else v
def nums(cells): return [num(c) for c in cells if num(c) is not None]
def rows(txt):
    out=[]
    for line in txt.split('\n'):
        if line.startswith('  | '):
            out.append([c.strip() for c in line[4:].split(' | ')])
        else: out.append(line)
    return out
def period_of(rcp, txt):
    m = re.search(r"(20\d\d)년\s*\(?'?\d*\.?\d*\.?\d*~?'?(\d\d)\.(\d\d)\.(\d\d)\)?", txt)
    return None
def read(d, key):
    fs=[f for f in glob.glob(os.path.join(d,'*.txt')) if key in os.path.basename(f)]
    return open(fs[0]).read() if fs else ''
res={}
for d in sorted(glob.glob(os.path.join(BASE,'2*/'))):
    rcp=os.path.basename(d.rstrip('/'))
    main=open(os.path.join(d,'main.html'),encoding='utf-8',errors='ignore').read()
    t=re.search(r'<title>(.*?)</title>',main,re.S)
    title=re.sub(r'\s+',' ',t.group(1)) if t else ''
    biz = read(d,'II_사업의_내용')
    prod = read(d,'주요_제품') or biz
    raw = read(d,'원재료') or biz
    sales = read(d,'매출_및_수주') or read(d,'매출에') or biz
    seg = read(d,'영업부문정보_연결') or read(d,'연결재무제표_주석')
    r={'title':title}
    # ASP sentence
    m=re.search(r'가격변동추이(.*?)(?:\n|당사는 국내 및 해외 원자재|$)',prod,re.S)
    r['asp_text']=re.sub(r'\s+',' ',m.group(1))[:400] if m else ''
    asp={}
    for nm,key in [('MLCC','MLCC'),('카메라모듈','CAM'),('반도체패키지기판','PKG'),('패키지기판','PKG')]:
        mm=re.search(nm.replace('MLCC','M ?LCC')+r'의\s*평균판매가격[은도는]*\s*전년\s*대비\s*([\d.]+)\s*%\s*(상승|하락)', r['asp_text'])
        if mm and key not in asp:
            v=float(mm.group(1)); asp[key]= v if mm.group(2)=='상승' else -v
    r['asp']=asp
    # raw material price
    m=re.search(r'주요 원재료 등의 가격변동추이\s*(.*?)(?:다\.|\n)',raw,re.S)
    r['rm_text']=re.sub(r'\s+',' ',m.group(1))[:400] if m else ''
    # capacity / production / utilization
    R=rows(raw)
    def unit_rows(after_kw, before_kw):
        out={}; on=False
        for x in R:
            if isinstance(x,str):
                if after_kw in x: on=True
                elif on and before_kw in x: break
                continue
            if on and x and re.search(r'(億個|억개|百万個|백만개|千㎡|천㎡)',x[0]):
                segname=re.sub(r'\s*\(.*','',x[0]).strip()
                k=SEGMAP.get(segname,segname)
                out[k]=nums(x[1:])
        return out
    r['capa']=unit_rows('생산능력 및','생산실적 및')
    r['prod']=unit_rows('(1) 생산실적','(2) 가동률')
    util={}
    on=False
    for x in R:
        if isinstance(x,str):
            if '(2) 가동률' in x: on=True
            continue
        if on and x and re.search(r'(億個|억개|百万個|백만개|千㎡|천㎡)',x[0]):
            segname=re.sub(r'\s*\(.*','',x[0]).strip(); k=SEGMAP.get(segname,segname)
            util[k]=x[-1]
        if on and x and x[0].startswith('사업부문') and util: break
    r['util']=util
    # capex by segment (억원)
    cap={}
    m=re.search(r'부문별 시설투자금액(.*?)(?:\n[^ ]|$)',raw,re.S)
    if m:
        for line in m.group(1).split('\n'):
            cells=[c.strip() for c in line.strip(' |').split(' | ')]
            if cells and cells[0] in SEGMAP or (cells and cells[0] in ('연구개발','합 계','합계')):
                v=nums(cells[-1:])
                if v: cap[SEGMAP.get(cells[0],cells[0])]=v[0]
    r['capex_eok']=cap
    # sales table (수출/내수/합계)
    S=rows(sales); st={}; cur=None; on=False
    for x in S:
        if isinstance(x,str):
            if '매출실적' in x: on=True
            if on and '매출 유형별' in x: break
            continue
        if not on or not x: continue
        if '수출' in x and x[0] not in ('수출',):
            name=x[0].replace(' ','')
            cur={'합계':'TOTAL'}.get(name, SEGMAP.get(x[0], x[0]))
            i=x.index('수출'); st.setdefault(cur,{})['수출']=nums(x[i+1:])
        elif x[0] in ('내수','합계') and cur:
            st[cur][x[0]]=nums(x[1:])
    r['sales']=st
    # segment note
    G=rows(seg); segs=[]; hdr=None; block=None
    started=False
    for x in G:
        if isinstance(x,str):
            if '영업부문' in x: started=True
            if started and ('지역에 대한' in x or '4. 현금' in x) and segs: started=False
            continue
        if not started and not any('영업부문' in c for c in x): 
            continue
        started=True
        cand=[SEGMAP.get(c) for c in x if c in SEGMAP]
        if len(cand)==3: hdr=cand; block={}; segs.append(block); continue
        if hdr is not None and x and (x[0].startswith('매출액') or x[0]=='수익' or x[0].startswith('수익(')):
            v=nums(x[1:]); block['rev']=dict(zip(hdr+['TOTAL'],v))
        if hdr is not None and x and x[0].startswith('영업이익'):
            v=nums(x[1:]); block['op']=dict(zip(hdr+['TOTAL'],v))
        if hdr is not None and x and x[0].startswith('감가상각비'):
            v=nums(x[1:]); block['dep']=dict(zip(hdr+['TOTAL'],v))
        if any('지역에 대한' in c for c in x) and len(segs)>=2: break
    r['seg']=[b for b in segs if b.get('rev')][:2]
    # region
    regs=[]; names=None
    for x in G:
        if isinstance(x,str): continue
        if '내수' in x and ('중국' in x or '중 국' in x):
            names=[c.replace(' ','') for c in x if c and c.replace(' ','') in ('내수','수출','중국','동남아','미주','유럽','일본','연결조정')]
        elif x and x[0]=='내 수' : pass
        elif x and (x[0].startswith('매출액') or x[0]=='수익' or x[0].startswith('수익(')) and names and len(nums(x[1:]))>=7:
            v=nums(x[1:]); 
            regs.append({'names':names,'vals':v})
    # older format: header '구 분 | 국 내 | 중 국 ...' and subheader '내 수 | 수 출'
    if not regs:
        for i,x in enumerate(G):
            if isinstance(x,list) and x and x[0].replace(' ','')=='구분' and any(c.replace(' ','')=='중국' for c in x):
                nm=['내수','수출','중국','동남아','미주','유럽','일본','연결조정']
                for y in G[i+1:i+4]:
                    if isinstance(y,list) and y and y[0].startswith('매출액'):
                        regs.append({'names':nm,'vals':nums(y[1:])}); break
    r['region']=regs[:2]
    # samsung electronics
    m=re.findall(r'삼성전자와 그 종속\S*[^\n]*?매출은\s*([\d,]+)\s*억원(?:\(전\S*:\s*([\d,]+)억원\))?',seg)
    m2=re.search(r'삼성전자와 그 종속\S*(?: 및 \S+)?이며,\s*당\S* 해당 거래처에 대한 매출은 각각\s*([\d,]+)억원\(전\S*:\s*([\d,]+)억원\)',seg)
    r['sec']= [m2.group(1),m2.group(2)] if m2 else ([m[0][0], m[1][0] if len(m)>1 else m[0][1]] if m else [])
    tb=re.findall(r'\|\s*수익(?:\(매출액\))?\s*\|\s*([\d,]+)\s*\n',seg)
    r['sec_tbl']=tb[:2]
    # 주요 매출처 sentence in sales section
    m=re.search(r'주요 매출처[^\n]*?약\s*([\d.]+)\s*%',sales)
    r['sec_pct']=m.group(1) if m else ''
    res[rcp]=r
json.dump(res,open(os.path.join(BASE,'parsed.json'),'w'),ensure_ascii=False,indent=1)
for k,v in res.items():
    print(k, v['title'][:30])
    print('  asp',v['asp'],'| util',v['util'],'| capex',v['capex_eok'])
    print('  capa',v['capa']); print('  prod',v['prod'])
    print('  sales',{a:b.get('합계') for a,b in v['sales'].items()})
    print('  seg', [ (b.get('rev'),b.get('op')) for b in v['seg']])
    print('  region', v['region'])
    print('  sec', v['sec'], v['sec_tbl'], v['sec_pct'])
