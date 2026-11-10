"""삼성전기 매출추정 워크북 빌더 (SMIC 양식).

사용: python scripts/build_workbook.py data/raw_data.json 삼성전기_매출추정.xlsx
- 모든 금액은 '원' 단위로 셀에 풀로 입력하고, 표시형식으로 백만원/십억원을 보여준다.
- 추정 가정(파란 글씨)만 바꾸면 매출추정 → REV 까지 수식으로 연동된다.
"""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment

D = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
M = 1_000_000

# ---------------------------------------------------------------- styles
RED, GRAY_D, GRAY_M, GRAY_L, YEL, HL_RED = 'C00000', '595959', '7F7F7F', 'D0CECE', 'F9F8BE', 'D94B4B'
PAST_FILL = PatternFill('solid', fgColor='F2F2F2')
EST_FILL = PatternFill('solid', fgColor=YEL)
def fnt(kind='L', size=10, color='000000', bold=False, italic=False):
    name = {'L': 'AppleSDGothicNeo L00', 'M': 'AppleSDGothicNeo M00', 'B': 'AppleSDGothicNeo B00',
            'H': 'AppleSDGothicNeo H00', 'N': 'Arial'}[kind]
    return Font(name=name, size=size, color=color, bold=bold, italic=italic)
thin = Side(style='thin', color='000000'); dbl = Side(style='double', color='000000'); thick = Side(style='medium', color='000000')
hair = Side(style='thin', color='BFBFBF')
F_WON = '#,##0,,;(#,##0,,);"-"'          # 원 -> 백만원
F_WON_B = '#,##0,,,;(#,##0,,,);"-"'      # 원 -> 십억원
F_USD = '#,##0.0,,;(#,##0.0,,);"-"'      # USD -> $mn
F_PCT = '0.0%;(0.0%);"-"'
F_NUM = '#,##0;(#,##0);"-"'
F_NUM1 = '#,##0.0;(#,##0.0);"-"'
F_NUM2 = '#,##0.00;(#,##0.00);"-"'
F_FX = '#,##0.0'

wb = Workbook()
wb.remove(wb.active)
SHEETS = {}

class S:
    """시트 래퍼: 표를 위에서 아래로 쌓고, (표키, 행키) -> 행번호, (표키, 열키) -> 열문자 를 기억한다."""
    def __init__(self, name, widths=None):
        self.ws = wb.create_sheet(name); self.name = name; self.r = 2; self.rows = {}; self.cols = {}
        SHEETS[name] = self
        self.ws.sheet_view.showGridLines = False
        self.ws.column_dimensions['A'].width = 2.5
        self.ws.column_dimensions['B'].width = (widths or {}).get('B', 34)
        for i in range(3, 40):
            self.ws.column_dimensions[CL(i)].width = (widths or {}).get('data', 12.5)
    def ref(self, tk, rk, ck, absolute=False, from_sheet=None):
        r = self.rows[(tk, rk)]; c = self.cols[(tk, ck)]
        a = f"${c}${r}" if absolute else f"{c}{r}"
        if from_sheet is not None and from_sheet != self.name:
            return f"'{self.name}'!{a}"
        return a
    def cell(self, tk, rk, ck): return self.ws[f"{self.cols[(tk, ck)]}{self.rows[(tk, rk)]}"]
    def text(self, txt, kind='L', size=10, color='000000', bold=False, col='B', indent=0, mark=None):
        c = self.ws[f"{col}{self.r}"]; c.value = txt; c.font = fnt(kind, size, color, bold)
        c.alignment = Alignment(indent=indent, vertical='top', wrap_text=False)
        if mark: self.ws[f"A{self.r}"] = mark; self.ws[f"A{self.r}"].font = fnt('B', 10)
        self.r += 1
    def gap(self, n=1): self.r += n

    def table(self, tk, title, unit, cols, rows, final=False, est_from=None, vlines=(), note=None):
        """cols: [(colkey, header)] ; rows: [dict(key,label,vals:{colkey:value|formula|callable}, style, fmt, indent)]
        est_from: 이 열키부터 추정(연노랑). vlines: 왼쪽에 세로줄을 넣을 열키 목록."""
        ws = self.ws
        ws[f"A{self.r}"] = 'X'; ws[f"A{self.r}"].font = fnt('B', 10)
        ws[f"B{self.r}"] = title; ws[f"B{self.r}"].font = fnt('H', 11, RED if final else '000000')
        self.r += 1
        hr = self.r
        hdr_fill = PatternFill('solid', fgColor=RED if final else GRAY_D)
        c = ws[f"B{hr}"]; c.value = unit; c.font = fnt('B', 9, 'FFFFFF'); c.fill = hdr_fill
        c.border = Border(bottom=dbl); c.alignment = Alignment(horizontal='left', vertical='center')
        est = False
        colmap = {}
        for j, (ck, h) in enumerate(cols):
            L = CL(3 + j); colmap[ck] = L; self.cols[(tk, ck)] = L
            x = ws[f"{L}{hr}"]; x.value = h; x.font = fnt('B', 9, 'FFFFFF'); x.fill = hdr_fill
            x.alignment = Alignment(horizontal='center', vertical='center'); x.border = Border(bottom=dbl)
        last_col = CL(2 + len(cols))
        self.r += 1
        first = self.r
        for i, row in enumerate(rows):
            rr = self.r; key = row.get('key', f"_r{i}"); self.rows[(tk, key)] = rr
            style = row.get('style', 'num'); fmt = row.get('fmt', F_WON)
            lab = ws[f"B{rr}"]; lab.value = row['label']
            ind = row.get('indent', 0)
            if style in ('yoy', 'pct', 'sub'): ind = max(ind, 1)
            lab.alignment = Alignment(indent=ind)
            lab.font = fnt('B' if style == 'bold' else 'L', 10, RED if style == 'yoy' else (GRAY_M if style == 'pct' else '000000'),
                           italic=(style == 'yoy'))
            est = False
            for ck, h in cols:
                L = colmap[ck]
                if est_from is not None and ck == est_from: est = True
                v = row.get('vals', {}).get(ck)
                if callable(v): v = v(ck)
                x = ws[f"{L}{rr}"]
                if v is not None: x.value = v
                inp = ck in row.get('inputs', ()) or row.get('input_all')
                color = '0000FF' if inp else (RED if style == 'yoy' else (GRAY_M if style == 'pct' else '000000'))
                x.font = Font(name='Arial', size=10, color=color, bold=(style in ('yoy', 'bold')), italic=(style == 'yoy'))
                x.number_format = row.get('fmts', {}).get(ck, fmt)
                x.alignment = Alignment(horizontal='right' if style not in ('text',) else 'left')
                if style == 'text': x.font = fnt('L', 9); x.number_format = '@'
                if est_from is not None: x.fill = EST_FILL if est else PAST_FILL
            if row.get('comment'):
                lab.comment = Comment(row['comment'], 'build')
            self.r += 1
        lastr = self.r - 1
        # 테두리: 마지막 행 굵은선, 세로줄
        for j in range(2, 3 + len(cols)):
            L = CL(j)
            x = ws[f"{L}{lastr}"]; b = x.border
            x.border = Border(left=b.left, right=b.right, top=b.top, bottom=thick)
        vl = ['__first__'] + list(vlines)
        for ck in vl:
            L = colmap[cols[0][0]] if ck == '__first__' else colmap[ck]
            for rr in range(hr, lastr + 1):
                x = ws[f"{L}{rr}"]; b = x.border
                x.border = Border(left=thin, right=b.right, top=b.top, bottom=b.bottom)
        if note:
            for n in (note if isinstance(note, list) else [note]):
                ws[f"B{self.r}"] = n; ws[f"B{self.r}"].font = fnt('L', 8, GRAY_M); self.r += 1
        self.r += 1
        return colmap

def yoy_formula(cur, prev):
    return f"=IFERROR({cur}/{prev}-1,\"\")"

# ================================================================ FX
fx = S('FX')
FXQ = D['fx_q']
qkeys = ['1Q23','2Q23','3Q23','4Q23','1Q24','2Q24','3Q24','4Q24','1Q25','2Q25','3Q25','4Q25','1Q26','2Q26','3Q26']
fx.text('원/달러 환율 (분기 평균) — 수출 비중 97%라 원화 매출의 핵심 변수', 'H', 11, RED, mark='X')
fx.text(D['fx_note'], 'L', 8, GRAY_M)
fx.gap()
cols = [(k, k) for k in qkeys] + [('4Q26QTD', '4Q26 QTD')]
fx.table('q', 'USD/KRW 분기 평균', '(원/달러)', cols,
         [dict(key='fx', label='USD/KRW 평균', fmt=F_FX, vals={k: FXQ[k] for k in qkeys} | {'4Q26QTD': FXQ['4Q26']}),
          dict(key='qoq', label='QoQ (%)', style='yoy', fmt=F_PCT,
               vals={k: (yoy_formula(f"{CL(3+i)}{{r}}", f"{CL(2+i)}{{r}}") if i > 0 else None) for i, k in enumerate(qkeys + ['4Q26QTD'])})],
         vlines=['1Q24', '1Q25', '1Q26'])
# fix {r} placeholders -> fx row
for i in range(1, len(qkeys) + 1):
    c = fx.ws[f"{CL(3+i)}{fx.rows[('q','qoq')]}"]; c.value = c.value.replace('{r}', str(fx.rows[('q', 'fx')]))
def fxq(k, sheet=None):  # 분기 평균 참조
    return fx.ref('q', 'fx', k, absolute=True, from_sheet=sheet)
acols = [('2023', '2023'), ('2024', '2024'), ('2025', '2025'), ('1H25', '1H25'), ('2H25', '2H25'), ('1H26', '1H26'),
         ('2H26E', '2H26E'), ('2026E', '2026E'), ('2027E', '2027E'), ('2028E', '2028E')]
def avg_q(qs): return '=AVERAGE(' + ','.join(fxq(q) for q in qs) + ')'
fx.table('a', '연간·반기 평균 및 추정 가정', '(원/달러)', [('4Q26E', '4Q26E')] + acols,
         [dict(key='fx', label='USD/KRW 평균', fmt=F_FX, inputs=('4Q26E', '2027E', '2028E'),
               vals={'4Q26E': 1340, '2023': avg_q(qkeys[0:4]), '2024': avg_q(qkeys[4:8]), '2025': avg_q(qkeys[8:12]),
                     '1H25': avg_q(['1Q25', '2Q25']), '2H25': avg_q(['3Q25', '4Q25']), '1H26': avg_q(['1Q26', '2Q26']),
                     '2027E': 1340, '2028E': 1340}),
          ], est_from='4Q26E', vlines=['2023', '2026E'])
c4 = fx.ref('a', 'fx', '4Q26E')
fx.ws[f"{fx.cols[('a','2H26E')]}{fx.rows[('a','fx')]}"] = f"=AVERAGE({fxq('3Q26')},{c4})"
fx.ws[f"{fx.cols[('a','2026E')]}{fx.rows[('a','fx')]}"] = f"=AVERAGE({fxq('1Q26')},{fxq('2Q26')},{fxq('3Q26')},{c4})"
for k in ('2023', '2024', '2025', '1H25', '2H25', '1H26'):
    fx.ws[f"{fx.cols[('a',k)]}{fx.rows[('a','fx')]}"].fill = PAST_FILL
def fxa(k, sheet=None): return fx.ref('a', 'fx', k, absolute=True, from_sheet=sheet)
fx.text('부가설명', 'B', 10, mark='-')
for t in ['4Q26E·2027E·2028E 환율은 1,340원 고정(4Q26 QTD 평균 1,340원, 11/10 종가 1,337원). 매크로 변수는 보수적으로 = 현재 수준 유지.',
          '참고: 2026년 공급계약 공시 적용환율 1,503.8(5/20) → 1,544.2(6/29) → 1,475.6(7/22) → 1,376.5(8/31) → 1,352.0(9/28) → 1,358.5(10/6) → 1,348.5(10/18) → 1,336.3(10/29). 3Q26부터 원화 강세 급진전.',
          '내수 매출 비중은 3~4%에 불과하므로 추정에서는 부문 매출 전체를 달러 연동으로 간주(단순화).']:
    fx.text(t, 'L', 9, indent=1)

# ================================================================ R_IS
ris = S('R_IS')
IA, IQ = D['is_annual'], D['is_quarter']
iq_keys = ['1Q24','2Q24','3Q24','4Q24','1Q25','2Q25','3Q25','4Q25','1Q26','2Q26','3Q26']
icols = [('2022', '2022'), ('2023', '2023'), ('2024', '2024'), ('2025', '2025')] + [(k, k) for k in iq_keys]
items = [('rev', '매출액', '매출액'), ('cogs', '매출원가', '매출원가'), ('gp', '매출총이익', None), ('sga', '판매비와관리비', '판매비와관리비'),
         ('op', '영업이익', '영업이익'), ('noi', '영업외손익', None), ('ebt', '법인세비용차감전순이익', '법인세비용차감전순이익'),
         ('tax', '법인세비용', '법인세비용'), ('cont', '계속영업이익', '계속영업당기순이익'), ('disc', '중단영업손익', '중단영업손익'),
         ('ni', '당기순이익', None), ('ni_p', '지배주주순이익', '지배주주순이익')]
rows = []
for key, lab, src in items:
    vals = {}
    for ck, _ in icols:
        srcd = IA.get(ck) if ck in IA else IQ.get(ck)
        if ck in ('4Q24', '4Q25'):
            continue
        if src and srcd and src in srcd:
            vals[ck] = srcd[src]
    rows.append(dict(key=key, label=lab, vals=vals, style='bold' if key in ('rev', 'op', 'ni') else 'num'))
    if key == 'rev': rows.append(dict(key='rev_yoy', label='YoY (%)', style='yoy', fmt=F_PCT, vals={}))
    if key == 'gp': rows.append(dict(key='gpm', label='GPM (%)', style='pct', fmt=F_PCT, vals={}))
    if key == 'op': rows.append(dict(key='opm', label='OPM (%)', style='pct', fmt=F_PCT, vals={}))
    if key == 'ni': rows.append(dict(key='npm', label='NPM (%)', style='pct', fmt=F_PCT, vals={}))
ris.table('is', '연결 손익계산서 (DART)', '(단위: 백만 원)', icols, rows, vlines=['1Q24', '1Q25', '1Q26'],
          note=[D['is_note'], '4Q24·4Q25 = 연간 − (1Q~3Q) 수식. 3Q26 매출원가·판관비·법인세·계속영업이익은 IR(억원 반올림)이라 합계에 소폭 오차 존재.'])
W = ris.ws
def R_(k): return ris.rows[('is', k)]
def C_(k): return ris.cols[('is', k)]
for ck, _ in icols:
    L = C_(ck)
    W[f"{L}{R_('gp')}"] = f"={L}{R_('rev')}-{L}{R_('cogs')}"
    W[f"{L}{R_('noi')}"] = f"={L}{R_('ebt')}-{L}{R_('op')}"
    W[f"{L}{R_('ni')}"] = f"={L}{R_('cont')}+{L}{R_('disc')}"
    W[f"{L}{R_('gpm')}"] = f"=IFERROR({L}{R_('gp')}/{L}{R_('rev')},\"\")"
    W[f"{L}{R_('opm')}"] = f"=IFERROR({L}{R_('op')}/{L}{R_('rev')},\"\")"
    W[f"{L}{R_('npm')}"] = f"=IFERROR({L}{R_('ni')}/{L}{R_('rev')},\"\")"
for q4, fy, qs in (('4Q24', '2024', ['1Q24', '2Q24', '3Q24']), ('4Q25', '2025', ['1Q25', '2Q25', '3Q25'])):
    for key, _, src in items:
        if src is None: continue
        W[f"{C_(q4)}{R_(key)}"] = f"={C_(fy)}{R_(key)}-" + '-'.join(f"{C_(q)}{R_(key)}" for q in qs)
# YoY
for a, b in (('2023', '2022'), ('2024', '2023'), ('2025', '2024')):
    W[f"{C_(a)}{R_('rev_yoy')}"] = yoy_formula(f"{C_(a)}{R_('rev')}", f"{C_(b)}{R_('rev')}")
for i, q in enumerate(iq_keys):
    if i >= 4: W[f"{C_(q)}{R_('rev_yoy')}"] = yoy_formula(f"{C_(q)}{R_('rev')}", f"{C_(iq_keys[i-4])}{R_('rev')}")
W[f"{C_('3Q26')}{R_('ni')}"].comment = Comment('잠정실적 공시 당기순이익 485,034백만원. IR 계속영업이익(억원 반올림) 기준으로 계산되어 소폭 차이.', 'build')

# ================================================================ R_REV
rr = S('R_REV')
SEG = [('COMP', '컴포넌트 (MLCC 등 수동소자)'), ('PKG', '패키지솔루션 (반도체 패키지기판)'), ('OPT', '광학솔루션 (카메라모듈)')]
cumk = ['1Q24','1H24','9M24','FY24','1Q25','1H25','9M25','FY25','1Q26','1H26']
C = D['seg_rev_cum']
rows = [dict(key=s, label=l, vals={k: C[k][s] for k in cumk}) for s, l in SEG]
rows.append(dict(key='TOTAL', label='합계', style='bold', vals={k: C[k]['TOTAL'] for k in cumk}))
rows.append(dict(key='chk', label='(check) 부문합 − 합계', style='pct', fmt=F_WON, vals={}))
rr.table('cum', '부문별 매출 — 누적 (DART 분기·반기·사업보고서 영업부문 주석)', '(단위: 백만 원)', [(k, k) for k in cumk], rows,
         vlines=['1Q25', '1Q26'], note='2024 분기 누적은 2025년 보고서의 전기(재작성) 컬럼 사용 → 통신모듈 중단영업 재분류가 반영된 동일 기준')
for k in cumk:
    L = rr.cols[('cum', k)]
    rr.ws[f"{L}{rr.rows[('cum','chk')]}"] = f"={L}{rr.rows[('cum','COMP')]}+{L}{rr.rows[('cum','PKG')]}+{L}{rr.rows[('cum','OPT')]}-{L}{rr.rows[('cum','TOTAL')]}"
def cumref(s, k): return rr.ref('cum', s, k)
QK = iq_keys
deriv = {'1Q24': ('1Q24', None), '2Q24': ('1H24', '1Q24'), '3Q24': ('9M24', '1H24'), '4Q24': ('FY24', '9M24'),
         '1Q25': ('1Q25', None), '2Q25': ('1H25', '1Q25'), '3Q25': ('9M25', '1H25'), '4Q25': ('FY25', '9M25'),
         '1Q26': ('1Q26', None), '2Q26': ('1H26', '1Q26')}
I3 = D['seg_rev_3Q26_IR']
def qformula(s):
    def f(q):
        if q == '3Q26':
            return I3[s] if s != 'TOTAL' else None
        a, b = deriv[q]
        return f"={cumref(s, a)}" + (f"-{cumref(s, b)}" if b else '')
    return f
rows = [dict(key=s, label=l, vals={q: qformula(s)(q) for q in QK}) for s, l in SEG]
rows.append(dict(key='TOTAL', label='합계', style='bold', vals={}))
rows.append(dict(key='TOTAL_prelim', label='(참고) 잠정공시 매출 합계', style='pct', vals={'3Q26': I3['TOTAL_prelim']}))
for s, l in SEG:
    rows.append(dict(key=s + '_yoy', label=f"{l.split(' ')[0]} YoY (%)", style='yoy', fmt=F_PCT, vals={}))
rows.append(dict(key='TOTAL_yoy', label='합계 YoY (%)', style='yoy', fmt=F_PCT, vals={}))
for s, l in SEG:
    rows.append(dict(key=s + '_qoq', label=f"{l.split(' ')[0]} QoQ (%)", style='pct', fmt=F_PCT, vals={}))
rr.table('q', '부문별 매출 — 분기 (3개월, 누적에서 차감)', '(단위: 백만 원)', [(k, k) for k in QK], rows, vlines=['1Q25', '1Q26'],
         note=['3Q26 부문 매출은 IR 실적발표(2026.10.28, 억원). 3Q26 분기보고서(제출기한 11/16) 공시 후 백만원 단위로 교체 필요.'])
for q in QK:
    L = rr.cols[('q', q)]
    rr.ws[f"{L}{rr.rows[('q','TOTAL')]}"] = f"=SUM({L}{rr.rows[('q','COMP')]}:{L}{rr.rows[('q','OPT')]})"
for i, q in enumerate(QK):
    L = rr.cols[('q', q)]
    for s in ('COMP', 'PKG', 'OPT', 'TOTAL'):
        if i >= 4:
            rr.ws[f"{L}{rr.rows[('q', s + '_yoy')]}"] = yoy_formula(f"{L}{rr.rows[('q', s)]}", f"{rr.cols[('q', QK[i-4])]}{rr.rows[('q', s)]}")
        if i >= 1 and s != 'TOTAL':
            rr.ws[f"{L}{rr.rows[('q', s + '_qoq')]}"] = yoy_formula(f"{L}{rr.rows[('q', s)]}", f"{rr.cols[('q', QK[i-1])]}{rr.rows[('q', s)]}")
def qref(s, q, sheet=None): return rr.ref('q', s, q, from_sheet=sheet)

# 연간
A = D['seg_rev_annual']; ak = ['2021', '2022', '2023', '2024', '2025']
rows = [dict(key=s, label=l, vals={k: A[k][s] for k in ak}) for s, l in SEG]
rows.append(dict(key='TOTAL', label='합계', style='bold', vals={k: A[k]['TOTAL'] for k in ak}))
for s, l in SEG + [('TOTAL', '합계')]:
    rows.append(dict(key=s + '_yoy', label=f"{l.split(' ')[0]} YoY (%)", style='yoy', fmt=F_PCT, vals={}))
for s, l in SEG:
    rows.append(dict(key=s + '_mix', label=f"{l.split(' ')[0]} 비중 (%)", style='pct', fmt=F_PCT, vals={}))
rr.table('a', '부문별 매출 — 연간 (DART 사업보고서 「매출 및 수주상황」)', '(단위: 백만 원)', [(k, k) for k in ak], rows,
         note=A['basis'])
for i, k in enumerate(ak):
    L = rr.cols[('a', k)]
    for s in ('COMP', 'PKG', 'OPT', 'TOTAL'):
        if i: rr.ws[f"{L}{rr.rows[('a', s + '_yoy')]}"] = yoy_formula(f"{L}{rr.rows[('a', s)]}", f"{rr.cols[('a', ak[i-1])]}{rr.rows[('a', s)]}")
        if s != 'TOTAL': rr.ws[f"{L}{rr.rows[('a', s + '_mix')]}"] = f"={L}{rr.rows[('a', s)]}/{L}{rr.rows[('a','TOTAL')]}"
def aref(s, k, sheet=None): return rr.ref('a', s, k, from_sheet=sheet)

# IR: Capacitor 비중
SH = D['ir_mlcc_share']; c23 = D['comp_2023_q']
mk = list(SH.keys())
mcols = [(k, k) for k in mk] + [('2023', '2023'), ('2024', '2024'), ('2025', '2025'), ('1-3Q26', '1~3Q26')]
rows = [dict(key='comp', label='컴포넌트 매출', vals={}),
        dict(key='share', label='Capacitors(MLCC) 비중 (IR)', style='pct', fmt=F_PCT, vals={k: SH[k] for k in mk}),
        dict(key='cap', label='Capacitors(MLCC) 매출', indent=1, vals={}),
        dict(key='ele', label='전자소자(인덕터·칩저항 등) 매출', indent=1, vals={})]
rr.table('ir', '컴포넌트 내 Capacitors(MLCC) 비중 — 회사 IR 실적발표 자료', '(단위: 백만 원)', mcols, rows,
         vlines=['1Q24', '1Q25', '1Q26', '2023'],
         note=['IR 슬라이드의 MLCC(2Q25부터 「Capacitors」로 명칭 변경) vs 전자소자 막대 비중. 2023년 분기 컴포넌트 매출은 DART 누적 차감(1Q23·1H23·9M23·FY23).',
               '연간 비중 = Σ(분기 매출×분기 비중) / 연간 매출'])
for k in mk:
    L = rr.cols[('ir', k)]
    rr.ws[f"{L}{rr.rows[('ir','comp')]}"] = c23[k] if k in c23 else f"={qref('COMP', k)}"
    rr.ws[f"{L}{rr.rows[('ir','cap')]}"] = f"={L}{rr.rows[('ir','comp')]}*{L}{rr.rows[('ir','share')]}"
    rr.ws[f"{L}{rr.rows[('ir','ele')]}"] = f"={L}{rr.rows[('ir','comp')]}-{L}{rr.rows[('ir','cap')]}"
grp = {'2023': mk[0:4], '2024': mk[4:8], '2025': mk[8:12], '1-3Q26': mk[12:15]}
for g, ks in grp.items():
    L = rr.cols[('ir', g)]; a, b = rr.cols[('ir', ks[0])], rr.cols[('ir', ks[-1])]
    for row in ('comp', 'cap', 'ele'):
        rr.ws[f"{L}{rr.rows[('ir',row)]}"] = f"=SUM({a}{rr.rows[('ir',row)]}:{b}{rr.rows[('ir',row)]})"
    rr.ws[f"{L}{rr.rows[('ir','share')]}"] = f"={L}{rr.rows[('ir','cap')]}/{L}{rr.rows[('ir','comp')]}"

# 대신증권 FCBGA/BGA
DS = D['daishin_pkg']
fk = ['1Q24', '2Q24', '3Q24', '4Q24', '1Q25', '2Q25', '3Q25', '4Q25', '1Q26', '2Q26', '3Q26']
rows = [dict(key='fc', label='FCBGA (증권사 추정)', vals={}), dict(key='bga', label='BGA (증권사 추정)', vals={}),
        dict(key='sum', label='합계', style='bold', vals={}),
        dict(key='fc_yoy', label='FCBGA YoY (%)', style='yoy', fmt=F_PCT, vals={}), dict(key='bga_yoy', label='BGA YoY (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='share', label='FCBGA 비중 (%)', style='pct', fmt=F_PCT, vals={}),
        dict(key='dart', label='(참고) DART 패키지솔루션 매출', style='pct', vals={})]
rr.table('fc', 'FCBGA / BGA 분해 — 회사 미공시, 대신증권 추정치 활용', '(단위: 백만 원)', [(k, k) for k in fk] + [('2024', '2024'), ('2025', '2025')], rows,
         vlines=['1Q25', '1Q26', '2024'],
         note=[DS['src'] + '. 2024 분기는 대신증권 2025년 분기별 YoY로 역산(=2025값/(1+YoY)).',
               '신한투자증권: 3Q26 기판 내 FC-BGA 비중 60~63% 추정 / IBK: 패키지 내 AI 비중 60% 수준 — 대신 추정(3Q26 62.5%)과 정합'])
yq = ['1Q', '2Q', '3Q', '4Q']
for i, k in enumerate(fk):
    L = rr.cols[('fc', k)]
    if k.endswith('24'):
        L25 = rr.cols[('fc', k.replace('24', '25'))]
        rr.ws[f"{L}{rr.rows[('fc','fc')]}"] = f"={L25}{rr.rows[('fc','fc')]}/(1+{DS['FCBGA_yoy_2025'][k[:2]]})"
        rr.ws[f"{L}{rr.rows[('fc','bga')]}"] = f"={L25}{rr.rows[('fc','bga')]}/(1+{DS['BGA_yoy_2025'][k[:2]]})"
    else:
        rr.ws[f"{L}{rr.rows[('fc','fc')]}"] = DS['FCBGA'][k] * 1e9
        rr.ws[f"{L}{rr.rows[('fc','bga')]}"] = DS['BGA'][k] * 1e9
    rr.ws[f"{L}{rr.rows[('fc','sum')]}"] = f"={L}{rr.rows[('fc','fc')]}+{L}{rr.rows[('fc','bga')]}"
    rr.ws[f"{L}{rr.rows[('fc','share')]}"] = f"={L}{rr.rows[('fc','fc')]}/{L}{rr.rows[('fc','sum')]}"
    rr.ws[f"{L}{rr.rows[('fc','dart')]}"] = f"={qref('PKG', k)}"
    if i >= 4:
        P4 = rr.cols[('fc', fk[i-4])]
        for row in ('fc', 'bga'):
            rr.ws[f"{L}{rr.rows[('fc', row + '_yoy')]}"] = yoy_formula(f"{L}{rr.rows[('fc', row)]}", f"{P4}{rr.rows[('fc', row)]}")
for g, ks in (('2024', fk[0:4]), ('2025', fk[4:8])):
    L = rr.cols[('fc', g)]; a, b = rr.cols[('fc', ks[0])], rr.cols[('fc', ks[-1])]
    for row in ('fc', 'bga', 'sum', 'dart'):
        rr.ws[f"{L}{rr.rows[('fc',row)]}"] = f"=SUM({a}{rr.rows[('fc',row)]}:{b}{rr.rows[('fc',row)]})"
    rr.ws[f"{L}{rr.rows[('fc','share')]}"] = f"={L}{rr.rows[('fc','fc')]}/{L}{rr.rows[('fc','sum')]}"
L24, L25 = rr.cols[('fc', '2024')], rr.cols[('fc', '2025')]
for row in ('fc', 'bga'):
    rr.ws[f"{L25}{rr.rows[('fc', row + '_yoy')]}"] = yoy_formula(f"{L25}{rr.rows[('fc', row)]}", f"{L24}{rr.rows[('fc', row)]}")

# 수출/내수
ED = D['export_domestic']; ek = ['2023', '2024', '2025', '1H25', '1H26']
rows = []
for s, l in SEG + [('TOTAL', '합계')]:
    rows.append(dict(key=s + '_x', label=f"{l.split(' ')[0]} 수출", vals={k: ED[k][s]['수출'] for k in ek}, style='bold' if s == 'TOTAL' else 'num'))
    rows.append(dict(key=s + '_d', label=f"{l.split(' ')[0]} 내수", indent=1, vals={k: ED[k][s]['내수'] for k in ek}))
rows.append(dict(key='xs', label='수출 비중 (%)', style='pct', fmt=F_PCT, vals={k: f"={{c}}{{x}}/({{c}}{{x}}+{{c}}{{d}})" for k in ek}))
rr.table('ex', '부문별 수출 / 내수 (DART 「매출 및 수주상황」)', '(단위: 백만 원)', [(k, k) for k in ek], rows, vlines=['1H25'])
for k in ek:
    L = rr.cols[('ex', k)]; c = rr.ws[f"{L}{rr.rows[('ex','xs')]}"]
    c.value = f"={L}{rr.rows[('ex','TOTAL_x')]}/({L}{rr.rows[('ex','TOTAL_x')]}+{L}{rr.rows[('ex','TOTAL_d')]})"

# 지역
RG = D['region']; rk = ['내수', '수출', '중국', '동남아', '미주', '유럽', '일본', '합계']
rows = [dict(key=n, label={'수출': '수출(국내사업장 로컬수출)'}.get(n, n), style='bold' if n == '합계' else 'num',
             vals={k: RG[k].get(n) for k in ek}) for n in rk]
rows += [dict(key=n + '_s', label=f"{n} 비중 (%)", style='pct', fmt=F_PCT, vals={}) for n in ['중국', '동남아', '미주']]
rr.table('rg', '지역별 매출 (DART 연결 주석 「지역에 대한 공시」, 법인 소재지 기준)', '(단위: 백만 원)', [(k, k) for k in ek], rows, vlines=['1H25'],
         note='지역 = 판매법인 소재지 기준. 국내 주요고객: 삼성전자·TSMC·Google / 중국·동남아: 삼성전자·Xiaomi·Tesla / 미주: Intel·Future·Bosch (2Q26 주석)')
for k in ek:
    L = rr.cols[('rg', k)]
    for n in ['중국', '동남아', '미주']:
        rr.ws[f"{L}{rr.rows[('rg', n + '_s')]}"] = f"={L}{rr.rows[('rg', n)]}/{L}{rr.rows[('rg','합계')]}"

# 삼성전자향
SEC = D['sec_rev']; sk = ['2022', '2023', '2024', '2025', '1H25', '1H26']
tot = {'2022': A['2022']['TOTAL'], '2023': A['2023']['TOTAL'], '2024': A['2024']['TOTAL'], '2025': A['2025']['TOTAL'],
       '1H25': C['1H25']['TOTAL'], '1H26': C['1H26']['TOTAL']}
rows = [dict(key='sec', label='삼성전자와 그 종속기업향 매출', vals={k: SEC[k] for k in sk}),
        dict(key='tot', label='연결 매출(중단영업 제외)', vals=tot),
        dict(key='s', label='비중 (%)', style='pct', fmt=F_PCT, vals={})]
rr.table('sec', '최대 고객 — 삼성전자향 매출 (DART 주석 「주요 고객에 대한 정보」)', '(단위: 백만 원)', [(k, k) for k in sk], rows, vlines=['1H25'],
         note='2022는 재작성 前 매출 기준 비중. 매출 10% 이상 고객은 삼성전자 1곳(2022년까지 Xiaomi 포함).')
for k in sk:
    L = rr.cols[('sec', k)]
    rr.ws[f"{L}{rr.rows[('sec','s')]}"] = f"={L}{rr.rows[('sec','sec')]}/{L}{rr.rows[('sec','tot')]}"

# 부문 영업이익 (분기)
OC = D['seg_op_cum']
rows = []
for s, l in SEG + [('TOTAL', '합계')]:
    rows.append(dict(key=s, label=f"{l.split(' ')[0]} 영업이익", style='bold' if s == 'TOTAL' else 'num', vals={}))
for s, l in SEG:
    rows.append(dict(key=s + '_m', label=f"{l.split(' ')[0]} OPM (%)", style='pct', fmt=F_PCT, vals={}))
opq = QK[:-1]
rr.table('op', '부문별 영업이익 — 분기 (DART 영업부문 주석, 누적 차감)', '(단위: 백만 원)', [('2023', '2023'), ('2024', '2024'), ('2025', '2025')] + [(k, k) for k in opq], rows,
         vlines=['1Q24', '1Q25', '1Q26'], note='3Q26 부문 영업이익은 IR 미공시(증권사 추정만 존재) → 분기보고서 공시 후 추가')
for s in ('COMP', 'PKG', 'OPT', 'TOTAL'):
    r_ = rr.rows[('op', s)]
    rr.ws[f"{rr.cols[('op','2023')]}{r_}"] = OC['FY23'][s]
    rr.ws[f"{rr.cols[('op','2024')]}{r_}"] = OC['FY24'][s]
    rr.ws[f"{rr.cols[('op','2025')]}{r_}"] = OC['FY25'][s]
    for q in opq:
        a, b = deriv[q]
        rr.ws[f"{rr.cols[('op', q)]}{r_}"] = OC[a][s] - (OC[b][s] if b else 0)
for ck in ['2023', '2024', '2025'] + opq:
    L = rr.cols[('op', ck)]
    for s in ('COMP', 'PKG', 'OPT'):
        revref = aref(s, ck) if ck in ('2023', '2024', '2025') else qref(s, ck)
        rr.ws[f"{L}{rr.rows[('op', s + '_m')]}"] = f"=IFERROR({L}{rr.rows[('op', s)]}/{revref},\"\")"

# ================================================================ R_CAPA
rc = S('R_CAPA')
CAP, PRD, UT = D['capa'], D['prod'], D['util_disclosed']
U = D['capa_units']
ck_ = ['2021', '2022', '2023', '2024', '2025', '1Q26', '1H26']
rows = []
for s, l in SEG:
    nm = l.split(' ')[0]
    rows.append(dict(key=s + '_c', label=f"{nm} 생산능력 ({U[s]})", style='bold', fmt=F_NUM, vals={k: CAP[k][s] for k in ck_}))
    rows.append(dict(key=s + '_p', label=f"{nm} 생산실적 ({U[s]})", fmt=F_NUM, indent=1, vals={k: PRD[k][s] for k in ck_}))
    rows.append(dict(key=s + '_u', label=f"{nm} 가동률 (%)", style='pct', fmt=F_PCT, vals={}))
    rows.append(dict(key=s + '_ud', label=f"{nm} 가동률 — 공시치", style='pct', fmt='@', vals={k: UT[k][s] for k in ck_ if k in UT}))
rc.table('a', '생산능력 · 생산실적 · 가동률 (DART 「원재료 및 생산설비」)', '(단위: 표기)', [(k, k) for k in ck_], rows, vlines=['1Q26'],
         note=['생산능력 = 시간당 평균 생산능력 × 1일 평균 가동시간 × 가동일수(글로벌 완제품 기준). 1H26 가동일 145일·24시간.',
               '수동소자 Capa(개수)가 2023 12,568억개 → 2025 11,159억개로 줄어든 것은 고용량·고적층(AI·전장) 비중 확대로 개당 공정시간이 늘었기 때문(대신증권 컨콜 노트) → Q(개수)보다 P(믹스)가 매출을 끈다.'])
for k in ck_:
    L = rc.cols[('a', k)]
    for s, _ in SEG:
        rc.ws[f"{L}{rc.rows[('a', s + '_u')]}"] = f"={L}{rc.rows[('a', s + '_p')]}/{L}{rc.rows[('a', s + '_c')]}"
def capref(s, kind, k, sheet=None): return rc.ref('a', f"{s}_{kind}", k, from_sheet=sheet)

# 누적 -> 반기
hk = ['1Q24', '1H24', '9M24', '2024', '1Q25', '1H25', '9M25', '2025', '1Q26', '1H26']
CAP['2024'], PRD['2024'] = CAP['2024'], PRD['2024']
rows = []
for s, l in SEG:
    nm = l.split(' ')[0]
    rows.append(dict(key=s + '_c', label=f"{nm} 생산능력 누적 ({U[s]})", fmt=F_NUM, vals={k: CAP[k][s] for k in hk}))
    rows.append(dict(key=s + '_p', label=f"{nm} 생산실적 누적 ({U[s]})", fmt=F_NUM, indent=1, vals={k: PRD[k][s] for k in hk}))
rc.table('cum', '생산능력 · 생산실적 — 분기보고서 누적치', '(단위: 표기)', [(k, k) for k in hk], rows, vlines=['1Q25', '1Q26'])
half = ['1H24', '2H24', '1H25', '2H25', '1H26']
rows = []
for s, l in SEG:
    nm = l.split(' ')[0]
    rows.append(dict(key=s + '_c', label=f"{nm} 생산능력 ({U[s]})", fmt=F_NUM, vals={}))
    rows.append(dict(key=s + '_p', label=f"{nm} 생산실적 ({U[s]})", fmt=F_NUM, indent=1, vals={}))
    rows.append(dict(key=s + '_u', label=f"{nm} 가동률 (%)", style='pct', fmt=F_PCT, vals={}))
rc.table('h', '생산능력 · 생산실적 — 반기 (누적 차감)', '(단위: 표기)', [(k, k) for k in half], rows, vlines=['1H25', '1H26'])
hsrc = {'1H24': ('1H24', None), '2H24': ('2024', '1H24'), '1H25': ('1H25', None), '2H25': ('2025', '1H25'), '1H26': ('1H26', None)}
for h, (a, b) in hsrc.items():
    L = rc.cols[('h', h)]
    for s, _ in SEG:
        for kind in ('c', 'p'):
            fa = rc.ref('cum', f"{s}_{kind}", a); fb = rc.ref('cum', f"{s}_{kind}", b) if b else None
            rc.ws[f"{L}{rc.rows[('h', f'{s}_{kind}')]}"] = f"={fa}" + (f"-{fb}" if fb else '')
        rc.ws[f"{L}{rc.rows[('h', s + '_u')]}"] = f"={L}{rc.rows[('h', s + '_p')]}/{L}{rc.rows[('h', s + '_c')]}"
def halfref(s, kind, h, sheet=None): return rc.ref('h', f"{s}_{kind}", h, from_sheet=sheet)

# Implied ASP
pk = ['2023', '2024', '2025', '1H25', '2H25', '1H26']
rows = []
unitmul = {'COMP': 1e8, 'PKG': 1e3, 'OPT': 1e6}   # 표기 단위 -> 개(㎡)
perlab = {'COMP': ('원/개', '$/천개', 1000), 'PKG': ('원/㎡', '$/㎡', 1), 'OPT': ('원/개', '$/개', 1)}
for s, l in SEG:
    nm = l.split(' ')[0]
    rows.append(dict(key=s + '_rev', label=f"{nm} 매출", vals={}))
    rows.append(dict(key=s + '_q', label=f"{nm} 생산량 ({U[s]})", fmt=F_NUM, indent=1, vals={}))
    rows.append(dict(key=s + '_krw', label=f"{nm} Implied ASP ({perlab[s][0]})", fmt=F_NUM2 if s == 'COMP' else F_NUM, indent=1, vals={}))
    rows.append(dict(key=s + '_usd', label=f"{nm} Implied ASP ({perlab[s][1]})", fmt=F_NUM2, indent=1, vals={}))
    rows.append(dict(key=s + '_yoy', label=f"{nm} USD ASP 변화 (%)", style='yoy', fmt=F_PCT, vals={}))
rc.table('asp', 'Implied ASP = 매출 ÷ 생산량 (DART 기반 P×Q 분해)', '(단위: 백만 원 / 표기)', [(k, k) for k in pk], rows, vlines=['1H25'],
         note=['생산량을 출하량 대용으로 사용(재고 변동 무시). 반기 USD ASP 변화는 직전 반기 대비, 연간은 전년 대비.', 'USD 환산 = 원화 매출 ÷ 해당기간 평균환율(FX 시트)'])
for k in pk:
    L = rc.cols[('asp', k)]
    for s, _ in SEG:
        if k in ('2023', '2024', '2025'):
            rev = aref(s, k, 'R_CAPA'); q = capref(s, 'p', k)
        else:
            q = halfref(s, 'p', k)
            if k == '1H25': rev = f"{rr.ref('cum', s, '1H25', from_sheet='R_CAPA')}"
            elif k == '2H25': rev = f"{rr.ref('cum', s, 'FY25', from_sheet='R_CAPA')}-{rr.ref('cum', s, '1H25', from_sheet='R_CAPA')}"
            else: rev = f"{rr.ref('cum', s, '1H26', from_sheet='R_CAPA')}"
        rc.ws[f"{L}{rc.rows[('asp', s + '_rev')]}"] = f"={rev}"
        rc.ws[f"{L}{rc.rows[('asp', s + '_q')]}"] = f"={q}"
        rc.ws[f"{L}{rc.rows[('asp', s + '_krw')]}"] = f"={L}{rc.rows[('asp', s + '_rev')]}/({L}{rc.rows[('asp', s + '_q')]}*{unitmul[s]})"
        rc.ws[f"{L}{rc.rows[('asp', s + '_usd')]}"] = f"={L}{rc.rows[('asp', s + '_krw')]}/{fxa(k, 'R_CAPA')}*{perlab[s][2]}"
for s, _ in SEG:
    for a, b in (('2024', '2023'), ('2025', '2024'), ('2H25', '1H25'), ('1H26', '2H25')):
        L = rc.cols[('asp', a)]; Lb = rc.cols[('asp', b)]
        rc.ws[f"{L}{rc.rows[('asp', s + '_yoy')]}"] = yoy_formula(f"{L}{rc.rows[('asp', s + '_usd')]}", f"{Lb}{rc.rows[('asp', s + '_usd')]}")

# 회사 공시 ASP 변동률 / 원재료
AY = D['asp_yoy']; ak2 = ['2022', '2023', '2024', '2025', '1Q26', '1H26']
rows = [dict(key='MLCC', label='MLCC 평균판매가격 (YoY)', style='yoy', fmt=F_PCT, vals={k: AY[k].get('MLCC', 0) / 100 for k in ak2}),
        dict(key='PKG', label='반도체패키지기판 평균판매가격 (YoY)', style='yoy', fmt=F_PCT, vals={k: (AY[k].get('PKG', 0) / 100) for k in ak2}),
        dict(key='CAM', label='카메라모듈 평균판매가격 (YoY)', style='yoy', fmt=F_PCT, vals={k: AY[k].get('CAM', 0) / 100 for k in ak2})]
rc.table('ay', '주요 제품 가격변동 — 회사 공시 (원화, 전년 동기 대비)', '(단위: %)', [(k, k) for k in ak2], rows, vlines=['1Q26'],
         note='출처: 각 보고서 「II-2. 주요 제품 및 서비스 — 나. 가격변동추이」. 1H26: MLCC +13.9%, 기판 +18.7%, 카메라 +2.4%')
RM = D['rm_yoy']; rk2 = ['2023', '2024', '2025', '1Q26', '1H26']
rows = [dict(key=n, label=f"{n} 평균 매입단가 (YoY)", style='yoy', fmt=F_PCT, vals={k: RM[k][n] / 100 for k in rk2}) for n in ['PASTE/POWDER', 'CCL/PPG', '센서IC']]
rc.table('rm', '주요 원재료 매입단가 변동 — 회사 공시', '(단위: %)', [(k, k) for k in rk2], rows, vlines=['1Q26'],
         note='PASTE/POWDER=MLCC 주재료(SHOEI 등) / CCL/PPG=기판 원판(Mitsubishi·Resonac) / 센서IC=카메라모듈(삼성전자·Sony)')
CX = D['capex_eok']; xk = ['2022', '2023', '2024', '2025', '1H25', '1H26']
def g(dct, *names):
    for n in names:
        if n in dct: return dct[n] * 1e8
rows = [dict(key='COMP', label='컴포넌트', vals={k: g(CX[k], 'COMP') for k in xk}),
        dict(key='PKG', label='패키지솔루션', vals={k: g(CX[k], 'PKG') for k in xk}),
        dict(key='OPT', label='광학솔루션', vals={k: g(CX[k], 'OPT') for k in xk}),
        dict(key='RD', label='연구개발 등', vals={k: g(CX[k], '연구개발') for k in xk}),
        dict(key='T', label='합계', style='bold', vals={k: g(CX[k], '합 계', '합계') for k in xk})]
rc.table('cx', '부문별 시설투자 (DART 「설비의 신설·매입 계획」)', '(단위: 십억 원)', [(k, k) for k in xk], rows, vlines=['1H25'],
         note='2026 반기 누적 컴포넌트 투자 5,208억원 = 2025 연간(7,219억)의 72% → MLCC 증설 가속. 기판은 9/28 세종 4.27조·베트남 2.51조 별도 이사회 결의(R_공시)')
for rr_ in rc.rows:
    pass
for k in xk:
    for s in ('COMP', 'PKG', 'OPT', 'RD', 'T'):
        rc.ws[f"{rc.cols[('cx', k)]}{rc.rows[('cx', s)]}"].number_format = F_WON_B

# ================================================================ R_공시
rg = S('R_공시', widths={'B': 30, 'data': 14})
CT = [('2026-05-20', '20260520800239', 'Silicon Capacitor 공급계약', 'SICAP', 1557034520000, 1035400000, 1503.80, '2027-01-01', '2028-12-31'),
      ('2026-06-30', '20260630800114', 'MLCC 공급계약', 'MLCC', 453994800000, 294000000, 1544.20, '2027-01-01', '2027-12-31'),
      ('2026-07-23', '20260723800073', 'MLCC 공급계약', 'MLCC', 295120000000, 200000000, 1475.60, '2027-01-01', '2027-12-31'),
      ('2026-09-01', '20260901800106', 'MLCC 공급계약', 'MLCC', 1072219807696, 778946464, 1376.50, '2027-01-01', '2027-12-31'),
      ('2026-09-29', '20260929800281', 'MLCC 및 Inductor 공급계약', 'MLCC', 285003046640, 210801070, 1352.00, '2027-01-01', '2027-12-31'),
      ('2026-10-06', '20261006800130', 'MLCC 공급계약', 'MLCC', 285611040000, 210240000, 1358.50, '2027-01-01', '2027-12-31'),
      ('2026-10-19', '20261019800610', '패키지기판 공급계약', 'PKG', 360487762500, 267325000, 1348.50, '2027-01-01', '2027-12-31'),
      ('2026-10-30', '20261030800604', 'MLCC 공급계약', 'MLCC', 879285400000, 658000000, 1336.30, '2026-12-21', '2027-12-21')]
ccols = [('date', '공시일'), ('rcp', '접수번호'), ('name', '계약명'), ('type', '구분'), ('krw', '계약금액(원)'), ('usd', '계약금액(USD)'),
         ('fx', '적용환율'), ('st', '시작일'), ('en', '종료일'), ('yrs', '기간(년)'), ('usd27', '2027 귀속 USD'), ('usd28', '2028 귀속 USD')]
rows = []
for i, (dt, rcp, nm, tp, krw, usd, f_, st, en) in enumerate(CT):
    yrs = 2 if tp == 'SICAP' else 1
    rows.append(dict(key=f"c{i}", label=f"{dt} {nm}", style='num',
                     vals={'date': dt, 'rcp': rcp, 'name': nm, 'type': tp, 'krw': krw, 'usd': usd, 'fx': f_, 'st': st, 'en': en, 'yrs': yrs},
                     fmts={'date': '@', 'rcp': '@', 'name': '@', 'type': '@', 'krw': F_WON_B, 'usd': F_USD, 'fx': F_FX, 'st': '@', 'en': '@', 'yrs': '0', 'usd27': F_USD, 'usd28': F_USD}))
rows.append(dict(key='mlcc', label='MLCC 계약 합계 (2027 납품)', style='bold', vals={}, fmts={'usd': F_USD, 'usd27': F_USD, 'krw': F_WON_B}))
rows.append(dict(key='sicap', label='실리콘 커패시터 합계', style='bold', vals={}, fmts={'usd': F_USD, 'usd27': F_USD, 'usd28': F_USD, 'krw': F_WON_B}))
rows.append(dict(key='pkg', label='패키지기판 합계', style='bold', vals={}, fmts={'usd': F_USD, 'usd27': F_USD, 'krw': F_WON_B}))
rg.table('ct', '2026년 단일판매·공급계약 공시 (장기공급계약, LTA) — 상대방 「글로벌 대형기업」, 경영상 비밀유지로 유보', '(단위: 십억 원 / $mn)', ccols, rows,
         note=['계약금액 USD는 공시 「기타 투자판단 참고사항」의 원문 금액. 2년 계약(실리콘 커패시터)은 연도별 균등 귀속 가정.',
               '회사 3Q26 컨콜: 「50여개 고객사와 장기공급계약 체결 완료 및 진행 중, 27년 (MLCC) 매출의 과반 이상이 LTA에서 발생」 → 공시된 건은 매출액 대비 일정 규모 이상만 해당(하한)'])
for i, (dt, rcp, nm, tp, krw, usd, f_, st, en) in enumerate(CT):
    r_ = rg.rows[('ct', f"c{i}")]
    rg.ws[f"{rg.cols[('ct','usd27')]}{r_}"] = f"={rg.cols[('ct','usd')]}{r_}/{rg.cols[('ct','yrs')]}{r_}"
    rg.ws[f"{rg.cols[('ct','usd28')]}{r_}"] = f"=IF({rg.cols[('ct','yrs')]}{r_}=2,{rg.cols[('ct','usd')]}{r_}/2,0)"
first, last = rg.rows[('ct', 'c0')], rg.rows[('ct', f"c{len(CT)-1}")]
T, Uc, U27, U28, K = rg.cols[('ct', 'type')], rg.cols[('ct', 'usd')], rg.cols[('ct', 'usd27')], rg.cols[('ct', 'usd28')], rg.cols[('ct', 'krw')]
for key, tp in (('mlcc', 'MLCC'), ('sicap', 'SICAP'), ('pkg', 'PKG')):
    r_ = rg.rows[('ct', key)]
    for col in (Uc, U27, U28, K):
        rg.ws[f"{col}{r_}"] = f"=SUMIF(${T}${first}:${T}${last},\"{tp}\",{col}{first}:{col}{last})"
def lta(key, col, sheet): return rg.ref('ct', key, col, absolute=True, from_sheet=sheet)
rows = [dict(key='s', label='신규시설투자 — 패키지기판 생산시설 증설(세종)', style='text', vals={'a': '4.27조원', 'b': '2026-09-28 ~ 2028-05-31', 'c': '20260928800810', 'd': '자기자본 대비 43.6%. 패키지기판 Capa 확대'}),
        dict(key='v', label='종속회사 유형자산취득 — 베트남 패키지기판 공장 증설', style='text', vals={'a': '2.51조원', 'b': '~2028-04-30', 'c': '20260928800821', 'd': 'SEMV(카메라모듈·반도체기판) 자산총액 대비 103%'}),
        dict(key='bs', label='장래사업·경영계획 — 부산 AI D/C 서버용 기판·MLCC', style='text', vals={'a': '약 15조원', 'b': '2026 ~ 2040', 'c': '20260703800494', 'd': '고성능 기판·고부가 MLCC 마더라인 핵심기지'}),
        dict(key='sj', label='장래사업·경영계획 — 세종 고성능 패키지기판 허브', style='text', vals={'a': '약 8조원', 'b': '2026 ~ 2040', 'c': '20260702800206', 'd': 'AI 서버용 패키지 설비 확충 + R&D'}),
        dict(key='gl', label='타법인주식취득 — 글라셈(유리기판 JV, 66.2%)', style='text', vals={'a': '3,191억원', 'b': '2026-11-05 납입', 'c': '20261105800383', 'd': '동우화인켐(스미토모화학 자회사)과 Glass Core 합작'}),
        dict(key='pr', label='연결 잠정실적(공정공시) 3Q26', style='text', vals={'a': '매출 3,763,278백만', 'b': 'OP 690,957백만', 'c': '20261028800169', 'd': 'QoQ +8.9% / YoY +30.3%'})]
rg.table('inv', '투자·사업계획 관련 공시 (2026)', '(구분)', [('a', '금액'), ('b', '기간'), ('c', '접수번호'), ('d', '비고')], rows,
         note='→ 기판 신규 Capa는 2028년 상반기 완공: 회사 컨콜 「27년부터 단계적 양산 시작, 28년부터 매출에 본격 기여」')
for r_ in range(rg.rows[('inv', 's')], rg.rows[('inv', 'pr')] + 1):
    for c_ in 'CDEF': rg.ws[f"{c_}{r_}"].alignment = Alignment(horizontal='left')
rg.ws.column_dimensions['F'].width = 40

# ================================================================ 컨센
cs = S('컨센', widths={'B': 30, 'data': 12.5})
cs.text('컨센서스 매출 — 출처별 정리', 'H', 11, RED, mark='X')
cs.text('FnGuide 컨센서스는 네이버페이증권(FnGuide 제공) 및 증권사 리포트 인용치, 증권사 추정은 3Q26 실적발표(10/28) 직후 리뷰 리포트 원문', 'L', 8, GRAY_M)
cs.gap()
rows = [dict(key='rev', label='매출액', style='bold', fmt=F_WON_B, vals={'2023': 8892.4e9, '2024': 10294.1e9, '2025': 11314.5e9, '2026E': 14413.8e9}),
        dict(key='op', label='영업이익', fmt=F_WON_B, vals={'2023': 660.5e9, '2024': 735.0e9, '2025': 913.3e9, '2026E': 2252.3e9}),
        dict(key='ni', label='당기순이익', fmt=F_WON_B, vals={'2023': 450.5e9, '2024': 703.2e9, '2025': 731.0e9, '2026E': 1783.1e9}),
        dict(key='opm', label='OPM (%)', style='pct', fmt=F_PCT, vals={})]
cs.table('fn', 'FnGuide 컨센서스 — 연간 (네이버페이증권, 2026-11-10 조회)', '(단위: 십억 원)', [('2023', '2023A'), ('2024', '2024A'), ('2025', '2025A'), ('2026E', '2026E')], rows,
         est_from='2026E', note='네이버 모바일 API(finance/annual)는 2026E까지만 제공. 2027E 이후는 아래 리포트 인용 컨센 및 증권사 개별 추정 사용.')
for k in ('2023', '2024', '2025', '2026E'):
    L = cs.cols[('fn', k)]; cs.ws[f"{L}{cs.rows[('fn','opm')]}"] = f"={L}{cs.rows[('fn','op')]}/{L}{cs.rows[('fn','rev')]}"
rows = [dict(key='rev', label='매출액', style='bold', fmt=F_WON_B, vals={'2Q25': 2784.6e9, '3Q25': 2889.0e9, '4Q25': 2902.1e9, '1Q26': 3209.1e9, '2Q26': 3457.2e9, '3Q26E': 3826.4e9, '3Q26A': 3763.3e9}),
        dict(key='op', label='영업이익', fmt=F_WON_B, vals={'2Q25': 213.0e9, '3Q25': 260.3e9, '4Q25': 239.5e9, '1Q26': 280.6e9, '2Q26': 440.4e9, '3Q26E': 645.3e9, '3Q26A': 691.0e9})]
cs.table('fnq', 'FnGuide 컨센서스 — 분기 (3Q26E는 실적발표 前 컨센)', '(단위: 십억 원)', [(k, k) for k in ['2Q25', '3Q25', '4Q25', '1Q26', '2Q26', '3Q26E', '3Q26A']], rows,
         note='3Q26 실적: 매출 3.76조(컨센 3.83조 대비 −1.7%, 원화 강세 영향), 영업이익 6,910억(컨센 6,453억 대비 +7.1%)')
rows = [dict(key='rev', label='매출액', style='bold', fmt=F_WON_B, vals={'2026E': 14449.8e9, '2027E': 18747.7e9}),
        dict(key='op', label='영업이익', fmt=F_WON_B, vals={'2026E': 2068.6e9, '2027E': 4162.7e9}),
        dict(key='ni', label='순이익', fmt=F_WON_B, vals={'2026E': 1695.7e9, '2027E': 3388.7e9})]
cs.table('fn2', 'FnGuide 컨센서스 — 2026-10-28 기준 (하나증권·신한투자증권 리포트 「Consensus Data」 인용)', '(단위: 십억 원)', [('2026E', '2026E'), ('2027E', '2027E')], rows,
         est_from='2026E', note='3Q26 실적발표 직전 컨센. 실적 이후 증권사들이 2027E를 일제히 상향(신한 +11.7%, 대신 +5.7%) → 현재 컨센은 이보다 높을 가능성')
BR = [('하나증권', '2026-10-29', 'BUY / TP 2,000,000', {'4Q26E': (2202.4, 982.2, 863.3, 4047.8), '2026E': (7143.3, 3417.8, 3916.4, 14477.4), '2027E': (11284.4, 4471.9, 3393.2, 19149.4), '2028E': (None, None, None, 26525.9)}),
      ('대신증권', '2026-10-29', 'BUY / TP 2,800,000', {'4Q26E': (2141, 1015, 773, 3928), '2026E': (7082, 3450, 3826, 14358), '2027E': (10347, 4694, 4094, 19135), '2028E': (None, None, None, 21440)}),
      ('미래에셋증권', '2026-10-29', 'BUY / TP 2,800,000', {'4Q26E': (2221.2, 993.3, 771.8, 3986.3), '2026E': (7162.1, 3428.9, 3824.9, 14415.9), '2027E': (10816.3, 4882.1, 3122.9, 18825.4), '2028E': (None, None, None, 24022)}),
      ('신한투자증권', '2026-10-29', 'BUY', {'4Q26E': (2227.7, 1000.3, 775.0, 4003.0), '2026E': (7168.6, 3436.0, 3828.0, 14432.6), '2027E': (11441.9, 4762.3, 3715.2, 19919.4), '2028E': (None, None, None, 26795.5)}),
      ('IBK투자증권', '2026-10-29', 'BUY / TP 1,900,000', {'4Q26E': (None, None, None, 3919.9), '2026E': (None, None, None, 14350), '2027E': (None, None, None, 18392), '2028E': (None, None, None, 21790)})]
pers = ['4Q26E', '2026E', '2027E', '2028E']
bcols = [(f"{p}_{x}", f"{p} {x}") for p in pers for x in ('컴포넌트', '패키지', '광학', '합계')]
rows = []
for nm, dt, op_, v in BR:
    vals = {}
    for p in pers:
        for j, x in enumerate(('컴포넌트', '패키지', '광학', '합계')):
            if v[p][j] is not None: vals[f"{p}_{x}"] = v[p][j] * 1e9
    rows.append(dict(key=nm, label=f"{nm} ({dt}, {op_})", fmt=F_WON_B, vals=vals))
rows.append(dict(key='avg', label='증권사 평균', style='bold', fmt=F_WON_B, vals={}))
rows.append(dict(key='max', label='최고', style='pct', fmt=F_WON_B, vals={}))
rows.append(dict(key='min', label='최저', style='pct', fmt=F_WON_B, vals={}))
cs.table('br', '증권사별 부문 매출 추정 (3Q26 실적 리뷰, 2026-10-29)', '(단위: 십억 원)', bcols, rows,
         vlines=['2026E_컴포넌트', '2027E_컴포넌트', '2028E_컴포넌트'],
         note=['하나 2027E 광학은 분기 합(3,393.2) 기준(표 합계란 4,090.5는 오기로 판단). IBK는 3Q26 컴포넌트 매출을 1.833조로 오기 → 부문 추정 제외, 합계만 사용.',
               '부문 2028E는 리포트 미제시(합계만). 신한 부문명: 모듈=광학, 기판=패키지.'])
f0, f1 = cs.rows[('br', BR[0][0])], cs.rows[('br', BR[-1][0])]
for ck, _ in bcols:
    L = cs.cols[('br', ck)]
    cs.ws[f"{L}{cs.rows[('br','avg')]}"] = f"=IFERROR(AVERAGE({L}{f0}:{L}{f1}),\"\")"
    cs.ws[f"{L}{cs.rows[('br','max')]}"] = f"=IF(COUNT({L}{f0}:{L}{f1})=0,\"\",MAX({L}{f0}:{L}{f1}))"
    cs.ws[f"{L}{cs.rows[('br','min')]}"] = f"=IF(COUNT({L}{f0}:{L}{f1})=0,\"\",MIN({L}{f0}:{L}{f1}))"
for c_ in range(3, 3 + len(bcols)): cs.ws.column_dimensions[CL(c_)].width = 11
cs.text('부가설명', 'B', 10, mark='-')
for t in ['증권사들은 3Q26 이후 「MLCC·FCBGA 공급부족 장기화 + ASP 상승」을 근거로 2027E 매출을 18.4~19.9조(+28~38% YoY)로 제시.',
          '2028E는 21.4~26.8조로 편차가 큼 → 신규 Capa(기판 세종·베트남, MLCC 필리핀·부산)의 매출 기여 시점과 가격 가정 차이.',
          '주요 컨콜 코멘트(10/28): MLCC 3Q blended ASP 두 자릿수 상승, 4Q도 큰 폭 성장 / 27년 MLCC 매출 과반이 LTA / 기판 신규 Capa 27년 단계 양산·28년 본격 기여 / 27년에도 타이트한 수급과 ASP 상승 지속.']:
    cs.text(t, 'L', 9, indent=1)
def cons(ck, sheet): return cs.ref('br', 'avg', ck, absolute=True, from_sheet=sheet)

# ================================================================ 매출추정 공통
EST_COLS = [('2023', '2023'), ('2024', '2024'), ('2025', '2025'), ('1Q26', '1Q26'), ('2Q26', '2Q26'), ('3Q26', '3Q26'),
            ('4Q26E', '4Q26E'), ('2026E', '2026E'), ('2027E', '2027E'), ('2028E', '2028E')]
QB_COLS = [(k, k) for k in ['1Q25', '2Q25', '3Q25', '4Q25', '1Q26', '2Q26', '3Q26']] + [('4Q26E', '4Q26E')]
PQ_COLS = [('2023', '2023'), ('2024', '2024'), ('2025', '2025'), ('1H26', '1H26'), ('2H26E', '2H26E'), ('2026E', '2026E'), ('2027E', '2027E'), ('2028E', '2028E')]
PREVQ = {'1Q26': '1Q25', '2Q26': '2Q25', '3Q26': '3Q25', '4Q26E': '4Q25'}

def segA_layout(sh, title, rows_extra, est_note):
    rows = [dict(key='rev', label=f"{title} 매출액", style='bold', vals={}),
            dict(key='yoy', label='YoY (%)', style='yoy', fmt=F_PCT, vals={})] + rows_extra
    sh.table('A', f"{title} 매출 추정 (→ 삼성전기_REV 연결)", '(단위: 백만 원)', EST_COLS, rows, final=True, est_from='4Q26E',
             vlines=['1Q26', '4Q26E', '2026E'], note=est_note)

def segA(sh, s, total_from):
    name = sh.name
    W = sh.ws; RA = lambda k: sh.rows[('A', k)]; CA = lambda k: sh.cols[('A', k)]
    for k in ('2023', '2024', '2025'):
        W[f"{CA(k)}{RA('rev')}"] = f"={aref(s, k, name)}"
    for k in ('1Q26', '2Q26', '3Q26'):
        W[f"{CA(k)}{RA('rev')}"] = f"={qref(s, k, name)}"
    W[f"{CA('4Q26E')}{RA('rev')}"] = f"={sh.ref('qb','krw','4Q26E')}"
    W[f"{CA('2026E')}{RA('rev')}"] = f"=SUM({CA('1Q26')}{RA('rev')}:{CA('4Q26E')}{RA('rev')})"
    for k in ('2027E', '2028E'):
        W[f"{CA(k)}{RA('rev')}"] = f"={sh.ref('pq', total_from, k)}"
    W[f"{CA('2024')}{RA('yoy')}"] = yoy_formula(f"{CA('2024')}{RA('rev')}", f"{CA('2023')}{RA('rev')}")
    W[f"{CA('2025')}{RA('yoy')}"] = yoy_formula(f"{CA('2025')}{RA('rev')}", f"{CA('2024')}{RA('rev')}")
    W[f"{CA('2023')}{RA('yoy')}"] = yoy_formula(f"{CA('2023')}{RA('rev')}", aref(s, '2022', name))
    for q, pq in PREVQ.items():
        W[f"{CA(q)}{RA('yoy')}"] = yoy_formula(f"{CA(q)}{RA('rev')}", qref(s, pq, name))
    W[f"{CA('2026E')}{RA('yoy')}"] = yoy_formula(f"{CA('2026E')}{RA('rev')}", f"{CA('2025')}{RA('rev')}")
    W[f"{CA('2027E')}{RA('yoy')}"] = yoy_formula(f"{CA('2027E')}{RA('rev')}", f"{CA('2026E')}{RA('rev')}")
    W[f"{CA('2028E')}{RA('yoy')}"] = yoy_formula(f"{CA('2028E')}{RA('rev')}", f"{CA('2027E')}{RA('rev')}")
    return RA, CA

def seg_est_sheet(name, s, title, unit_lab, unit_mul, asp_div, inputs, a_rows_extra, pq_extra, a_title=None, a_note=None):
    """공통: Block B(분기 USD) + Block C(P×Q) + Block A(최종) 구성.
    unit_mul: 표기단위->개(㎡) ; asp_div: ASP 표시 단위 (MLCC $/천개 -> 1000)"""
    sh = S(name)
    sh.text(title, 'H', 12, RED, mark='X')
    sh.gap()
    segA_layout(sh, a_title, a_rows_extra, a_note)
    # --- 가정 블록
    rows = [dict(key=k, label=lab, fmt=f, vals={'v': v}, inputs=('v',), comment=cm) for k, lab, v, f, cm in inputs]
    sh.table('in', '추정 가정 (파란 글씨 = 입력값, 바꾸면 전체 연동)', '(가정)', [('v', '값')], rows)
    IN = lambda k: sh.ref('in', k, 'v', absolute=True)
    # --- Block B: 분기 USD
    rows = [dict(key='krw', label='매출액 (원화)', vals={}),
            dict(key='fx', label='평균환율 (원/달러)', fmt=F_FX, vals={}),
            dict(key='usd', label='매출액 (USD, $mn)', fmt=F_USD, vals={}),
            dict(key='usd_qoq', label='USD QoQ (%)', style='yoy', fmt=F_PCT, vals={}),
            dict(key='krw_qoq', label='원화 QoQ (%)', style='pct', fmt=F_PCT, vals={})]
    sh.table('qb', '분기 달러 매출 — 원화 매출에서 환율 효과를 걷어낸 실질 성장', '(단위: 백만 원 / $mn)', QB_COLS, rows, est_from='4Q26E', vlines=['1Q26', '4Q26E'])
    for i, (q, _) in enumerate(QB_COLS):
        L = sh.cols[('qb', q)]
        if q == '4Q26E':
            prv = sh.cols[('qb', QB_COLS[i-1][0])]
            sh.ws[f"{L}{sh.rows[('qb','usd')]}"] = f"={prv}{sh.rows[('qb','usd')]}*(1+{IN('q4q')})*(1+{IN('q4p')})"
            sh.ws[f"{L}{sh.rows[('qb','fx')]}"] = f"={fxa('4Q26E', name)}"
            sh.ws[f"{L}{sh.rows[('qb','krw')]}"] = f"={L}{sh.rows[('qb','usd')]}*{L}{sh.rows[('qb','fx')]}"
        else:
            sh.ws[f"{L}{sh.rows[('qb','krw')]}"] = f"={qref(s, q, name)}"
            sh.ws[f"{L}{sh.rows[('qb','fx')]}"] = f"={fxq(q, name)}"
            sh.ws[f"{L}{sh.rows[('qb','usd')]}"] = f"={L}{sh.rows[('qb','krw')]}/{L}{sh.rows[('qb','fx')]}"
        if i:
            prv = sh.cols[('qb', QB_COLS[i-1][0])]
            for r_ in ('usd', 'krw'):
                sh.ws[f"{L}{sh.rows[('qb', r_ + '_qoq')]}"] = yoy_formula(f"{L}{sh.rows[('qb', r_)]}", f"{prv}{sh.rows[('qb', r_)]}")
    QB = lambda r_, q: sh.ref('qb', r_, q)
    # --- Block C: P x Q
    rows = [dict(key='capa', label=f"생산능력 ({unit_lab})", fmt=F_NUM, vals={}),
            dict(key='util', label='가동률 (%)', style='pct', fmt=F_PCT, vals={}),
            dict(key='q', label=f"Q: 생산량 ({unit_lab})", style='bold', fmt=F_NUM, vals={}),
            dict(key='q_yoy', label='Q 증감 (%)', style='yoy', fmt=F_PCT, vals={}),
            dict(key='asp', label=f"P: USD ASP ({ {1000:'$/천개',1:'$/개' if s=='OPT' else '$/㎡'}[asp_div] })", style='bold', fmt=F_NUM2, vals={}),
            dict(key='asp_yoy', label='P 증감 (%)', style='yoy', fmt=F_PCT, vals={}),
            dict(key='usd', label='USD 매출 = Q × P ($mn)', fmt=F_USD, vals={}),
            dict(key='fx', label='× 평균환율 (원/달러)', fmt=F_FX, vals={}),
            dict(key='krw', label='원화 매출 (P×Q×FX)', style='bold', vals={})] + pq_extra
    sh.table('pq', 'P × Q × FX 추정 (DART 생산능력·생산실적으로 Q를, 매출÷생산량으로 P를 분해)', '(단위: 백만 원 / 표기)', PQ_COLS, rows,
             est_from='2H26E', vlines=['1H26', '2026E'],
             note=['2023~1H26: 실적(R_CAPA·R_REV) / 2H26E: 3Q26 실적+4Q26E(분기 USD 블록) / 2027E·2028E: Q·P 가정으로 산출',
                   'P 증감: 2H26E는 1H26 대비, 2027E는 「2H26E(연말 출구 수준)」 대비, 2028E는 2027E 대비'])
    W = sh.ws; R = lambda k: sh.rows[('pq', k)]; Cc = lambda k: sh.cols[('pq', k)]
    for k in ('2023', '2024', '2025'):
        L = Cc(k)
        W[f"{L}{R('capa')}"] = f"={capref(s, 'c', k, name)}"
        W[f"{L}{R('q')}"] = f"={capref(s, 'p', k, name)}"
        W[f"{L}{R('krw')}"] = f"={aref(s, k, name)}"
        W[f"{L}{R('fx')}"] = f"={fxa(k, name)}"
    L = Cc('1H26')
    W[f"{L}{R('capa')}"] = f"={halfref(s, 'c', '1H26', name)}"
    W[f"{L}{R('q')}"] = f"={halfref(s, 'p', '1H26', name)}"
    W[f"{L}{R('krw')}"] = f"={QB('krw','1Q26')}+{QB('krw','2Q26')}"
    W[f"{L}{R('fx')}"] = f"={fxa('1H26', name)}"
    L = Cc('2H26E')
    W[f"{L}{R('capa')}"] = f"={Cc('1H26')}{R('capa')}*(1+{IN('h2capa')})"
    W[f"{L}{R('util')}"] = f"={IN('h2util')}"
    W[f"{L}{R('q')}"] = f"={L}{R('capa')}*{L}{R('util')}"
    W[f"{L}{R('usd')}"] = f"={QB('usd','3Q26')}+{QB('usd','4Q26E')}"
    W[f"{L}{R('krw')}"] = f"={QB('krw','3Q26')}+{QB('krw','4Q26E')}"
    W[f"{L}{R('fx')}"] = f"={L}{R('krw')}/{L}{R('usd')}"
    L = Cc('2026E')
    for r_ in ('capa', 'q', 'usd', 'krw'):
        W[f"{L}{R(r_)}"] = f"={Cc('1H26')}{R(r_)}+{Cc('2H26E')}{R(r_)}"
    W[f"{L}{R('fx')}"] = f"={L}{R('krw')}/{L}{R('usd')}"
    for k in ('2023', '2024', '2025', '1H26', '2026E'):
        L = Cc(k)
        if k != '2026E': W[f"{L}{R('usd')}"] = f"={L}{R('krw')}/{L}{R('fx')}"
    for k in ('2023', '2024', '2025', '1H26', '2H26E', '2026E'):
        L = Cc(k)
        if k != '2H26E': W[f"{L}{R('util')}"] = f"={L}{R('q')}/{L}{R('capa')}"
        W[f"{L}{R('asp')}"] = f"={L}{R('usd')}/({L}{R('q')}*{unit_mul})*{asp_div}"
    # 2027E/2028E
    L7, L8 = Cc('2027E'), Cc('2028E')
    W[f"{L7}{R('q')}"] = f"={Cc('2026E')}{R('q')}*(1+{IN('q27')})"
    W[f"{L7}{R('asp')}"] = f"={Cc('2H26E')}{R('asp')}*(1+{IN('p27')})"
    W[f"{L8}{R('asp')}"] = f"={L7}{R('asp')}*(1+{IN('p28')})"
    W[f"{L8}{R('q')}"] = f"={L7}{R('q')}*(1+{IN('q28')})"
    for L in (L7, L8):
        W[f"{L}{R('usd')}"] = f"={L}{R('q')}*{unit_mul}*{L}{R('asp')}/{asp_div}"
        W[f"{L}{R('fx')}"] = f"={fxa('2027E' if L == L7 else '2028E', name)}"
        W[f"{L}{R('krw')}"] = f"={L}{R('usd')}*{L}{R('fx')}"
    # 증감
    pairs_q = [('2024', '2023'), ('2025', '2024'), ('2026E', '2025'), ('2027E', '2026E'), ('2028E', '2027E')]
    for a, b in pairs_q:
        W[f"{Cc(a)}{R('q_yoy')}"] = yoy_formula(f"{Cc(a)}{R('q')}", f"{Cc(b)}{R('q')}")
    pairs_p = [('2024', '2023'), ('2025', '2024'), ('2H26E', '1H26'), ('2026E', '2025'), ('2027E', '2H26E'), ('2028E', '2027E')]
    for a, b in pairs_p:
        W[f"{Cc(a)}{R('asp_yoy')}"] = yoy_formula(f"{Cc(a)}{R('asp')}", f"{Cc(b)}{R('asp')}")
    return sh, IN, QB

# ================================================================ 컴포넌트
inputs = [('q4q', '4Q26E 생산·출하(Q) QoQ', 0.05, F_PCT, '컨콜: 4Q는 통상 IT 재고조정으로 정체되나 올해는 전장·AI서버 선단품 강세로 큰 폭 성장. 가동률 90%대라 Q 여력은 제한적'),
          ('q4p', '4Q26E USD ASP QoQ', 0.10, F_PCT, '3Q26 blended ASP 두 자릿수 상승(컨콜), 4Q도 믹스·ASP 상승 지속. 3Q 달러매출 QoQ +21%에 비해 보수적으로 합계 +15.5%'),
          ('h2capa', '2H26 생산능력 증가 (1H26 대비)', 0.03, F_PCT, '1H26 컴포넌트 투자 5,208억(2025 연간의 72%). 다만 고적층 믹스로 개수 기준 Capa는 잘 안 늘어남'),
          ('h2util', '2H26 가동률', 0.92, F_PCT, '1Q26 95% / 1H26 91% / 신한: 2026년 95% 추정 → 92%'),
          ('q27', '2027E Q 증가 (2026E 대비)', 0.06, F_PCT, '필리핀 신공장 2027년 양산 개시(최선단 고부가), 부산 마더라인. 단 고용량·고적층 비중 확대로 개수 증가는 제한'),
          ('p27', '2027E USD ASP 증가 (2H26E 대비)', 0.12, F_PCT, 'LTA 기반 가격(27년 매출 과반), 경쟁사(무라타 등) 저수익 범용품 단종 → IT·전장 가격 인상 가능성. 2H26 평균 대비 +12%는 4Q26 출구 대비 +7% 수준'),
          ('q28', '2028E Q 증가', 0.08, F_PCT, '필리핀 풀가동 + 부산 증설 반영'),
          ('p28', '2028E USD ASP 증가', 0.05, F_PCT, '공급부족 지속되나 무라타·태양유전·중화권 증설로 가격 상승폭 둔화'),
          ('share', 'Capacitors(MLCC) 비중 — 4Q26E 이후', 0.92, F_PCT, 'IR 3Q26 92% 유지(AI용 MLCC 비중 확대로 상승 여지 있으나 보수적으로 고정)'),
          ('dc3q', '(참고) 3Q26 컴포넌트 내 데이터센터 비중', 0.30, F_PCT, 'IBK투자증권 2026-10-29 (회사 미공시)')]
pq_extra = [dict(key='si_usd', label='(+) 실리콘 커패시터 LTA 매출 ($mn)', fmt=F_USD, vals={}),
            dict(key='si_krw', label='(+) 실리콘 커패시터 원화 매출', vals={}),
            dict(key='total', label='컴포넌트 매출 합계', style='bold', vals={})]
rows_extra = [dict(key='cap', label='Capacitors(MLCC)', indent=1, vals={}),
              dict(key='share', label='Capacitors 비중 (ex 실리콘커패시터)', style='pct', fmt=F_PCT, vals={}),
              dict(key='ele', label='전자소자 (인덕터·칩저항 등)', indent=1, vals={}),
              dict(key='si', label='실리콘 커패시터 (신규, LTA)', indent=1, vals={})]
cp_, CIN, CQB = seg_est_sheet('매출추정_컴포넌트', 'COMP', '컴포넌트(수동소자) 매출 추정 — Q(억개) × P(USD/천개) × 환율 + 실리콘 커패시터 LTA',
                              '억개', 1e8, 1000, inputs, rows_extra, pq_extra, '컴포넌트',
                              ['Capacitors 비중: 실적은 IR 분기 비중(R_REV), 4Q26E 이후는 가정값. 실리콘 커패시터는 2027~28 LTA($1,035.4mn, 2년 균등)만 반영.'])
W = cp_.ws; R = lambda k: cp_.rows[('pq', k)]; Cc = lambda k: cp_.cols[('pq', k)]
for k, _ in PQ_COLS:
    L = Cc(k)
    if k in ('2027E', '2028E'):
        W[f"{L}{R('si_usd')}"] = f"='R_공시'!${rg.cols[('ct', 'usd27' if k == '2027E' else 'usd28')]}${rg.rows[('ct','sicap')]}"
        W[f"{L}{R('si_krw')}"] = f"={L}{R('si_usd')}*{L}{R('fx')}"
    else:
        W[f"{L}{R('si_usd')}"] = 0; W[f"{L}{R('si_krw')}"] = 0
    W[f"{L}{R('total')}"] = f"={L}{R('krw')}+{L}{R('si_krw')}"
# Block D: cross check
rows = [dict(key='lta', label='공시된 MLCC LTA 2027 귀속 ($mn)', fmt=F_USD, vals={'v': f"='R_공시'!${rg.cols[('ct','usd27')]}${rg.rows[('ct','mlcc')]}"}),
        dict(key='cap27', label='당사 2027E Capacitors(MLCC) USD 매출 ($mn)', fmt=F_USD, vals={'v': f"={Cc('2027E')}{R('usd')}*{CIN('share')}"}),
        dict(key='cov', label='공시 LTA 커버리지 (%)', style='pct', fmt=F_PCT, vals={'v': '={c}{a}/{c}{b}'}),
        dict(key='cons', label='증권사 평균 2027E 컴포넌트 (원)', vals={'v': f"={cons('2027E_컴포넌트', '매출추정_컴포넌트')}"}),
        dict(key='cons_usd', label='→ 실리콘커패시터 제외 USD ($mn, 당사 환율 가정)', fmt=F_USD, vals={}),
        dict(key='cons_asp', label='→ 컨센이 함의하는 2027E ASP ($/천개, 당사 Q 가정)', fmt=F_NUM2, vals={}),
        dict(key='cons_p', label='→ 컨센 함의 ASP 증가율 (2H26E 대비)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='dc', label='(참고) 3Q26 데이터센터향 컴포넌트 매출 (원)', vals={})]
cp_.table('chk', '크로스체크 — LTA 커버리지 · 컨센 역산', '(체크)', [('v', '값')], rows,
          note=['회사: 「27년 MLCC 매출의 과반 이상이 LTA」 → 공시분(대형 계약)만으로 당사 추정의 약 1/3 커버, 나머지는 미공시 소형 LTA(50여 고객) + 스팟',
                '컨센 역산: 증권사 평균 2027E 컴포넌트를 당사 환율·Q로 맞추려면 ASP가 얼마나 올라야 하는지 → 당사 +12% 가정과의 차이가 곧 컨센과의 갭'])
Wc = cp_.ws; V = cp_.cols[('chk', 'v')]; rcx = lambda k: cp_.rows[('chk', k)]
Wc[f"{V}{rcx('cov')}"] = f"={V}{rcx('lta')}/{V}{rcx('cap27')}"
Wc[f"{V}{rcx('cons_usd')}"] = f"={V}{rcx('cons')}/{fxa('2027E', '매출추정_컴포넌트')}-{Cc('2027E')}{R('si_usd')}"
Wc[f"{V}{rcx('cons_asp')}"] = f"={V}{rcx('cons_usd')}/({Cc('2027E')}{R('q')}*1e8)*1000"
Wc[f"{V}{rcx('cons_p')}"] = f"={V}{rcx('cons_asp')}/{Cc('2H26E')}{R('asp')}-1"
Wc[f"{V}{rcx('dc')}"] = f"={CQB('krw','3Q26')}*{CIN('dc3q')}"

def block_a_hist(sh, s, rows_extra_fn):
    pass


RA, CA = segA(cp_, 'COMP', 'total')
W = cp_.ws
irmap = {'2023': '2023', '2024': '2024', '2025': '2025', '1Q26': '1Q26', '2Q26': '2Q26', '3Q26': '3Q26'}
for k, _ in EST_COLS:
    L = CA(k)
    if k in irmap:
        W[f"{L}{RA('share')}"] = f"={rr.ref('ir', 'share', irmap[k], from_sheet='매출추정_컴포넌트')}"
    elif k == '2026E':
        W[f"{L}{RA('share')}"] = f"=SUMPRODUCT({CA('1Q26')}{RA('share')}:{CA('4Q26E')}{RA('share')},{CA('1Q26')}{RA('rev')}:{CA('4Q26E')}{RA('rev')})/{L}{RA('rev')}"
    else:
        W[f"{L}{RA('share')}"] = f"={CIN('share')}"
    if k in ('2027E', '2028E'):
        W[f"{L}{RA('si')}"] = f"={cp_.ref('pq', 'si_krw', k)}"
    else:
        W[f"{L}{RA('si')}"] = 0
    W[f"{L}{RA('cap')}"] = f"=({L}{RA('rev')}-{L}{RA('si')})*{L}{RA('share')}"
    W[f"{L}{RA('ele')}"] = f"={L}{RA('rev')}-{L}{RA('si')}-{L}{RA('cap')}"
cp_.text('부가설명 — 컴포넌트', 'B', 10, mark='-')
for t in ['① 왜 이렇게 쪼갰나: DART는 「수동소자」 한 줄. IR이 Capacitors(MLCC) vs 전자소자 비중을 매 분기 공개(89~92%) → MLCC를 분리. 실리콘 커패시터는 LTA($1,035mn, 27~28년)로 매출이 「계약」으로 보이는 신규 품목이라 별도 라인.',
          '② Q(억개): DART 생산능력×가동률. 2023→2025 Capa는 12,568→11,159억개로 「감소」, 가동률은 70%→93%. 개수 Capa가 줄어드는 건 고용량·고적층(AI·전장) 품목이 개당 공정시간을 더 먹기 때문 → 성장의 본체는 Q가 아니라 P(믹스).',
          '③ P(USD/천개): 매출÷생산량. 2025 $3.5 → 1H26 $3.7 → 2H26E $4.9. 원화 기준 공시 ASP도 1H26 +13.9% YoY. 3Q26 달러 매출 QoQ +21%(원화 +14%, 환율 −5.5%).',
          '④ 환율: 수출 97%. 3Q26 평균 1,420원 → 4Q26 QTD 1,340원(−5.6%). 2027~28E 1,340원 고정 → 같은 달러 성장에도 원화 매출은 2026E 대비 낮게 잡힘(컨센과 갭의 큰 원인).',
          '⑤ 2027E: Q +6%(필리핀 신공장, 고적층 믹스 감안) × P +12%(2H26 대비; LTA 가격·범용품 가격 인상) → 달러 매출 +40% YoY. 2028E: Q +8%, P +5%.',
          '⑥ AI(데이터센터) 크로스체크: 3Q26 컴포넌트의 30%(IBK) ≈ 0.56조/분기. GPU당 MLCC 용량 매년 2배(회사 「AI MLCC 법칙」), 1005 47μF 27년 수요 8배 → 성장 방향은 확실하나 개수 Capa가 병목 → Q가 아니라 P로 반영.',
          '⑦ 리스크(하방): 환율 추가 하락, IT 수요(메모리 가격 급등으로 스마트폰·PC 출하 둔화) / (상방): LTA 추가 공시, 2027 범용 MLCC 가격 인상 본격화']:
    cp_.text(t, 'L', 9, indent=1)

# ================================================================ 패키지
inputs = [('q4q', '4Q26E 생산면적(Q) QoQ', 0.02, F_PCT, '1H26 가동률 89%, 공급부족으로 풀가동 근접 → 면적 증가는 제한'),
          ('q4p', '4Q26E USD ASP QoQ', 0.10, F_PCT, 'FCBGA 대면적·고다층 AI가속기/서버CPU 신제품 양산, 임베딩 기판 4Q 물량 2배(컨콜). 3Q 달러매출 QoQ +29% 대비 보수적으로 합계 +12%'),
          ('h2capa', '2H26 생산능력 증가 (1H26 대비)', 0.02, F_PCT, '신규 공장은 2028년 완공 → 2H26은 디보틀넥 수준'),
          ('h2util', '2H26 가동률', 0.93, F_PCT, '1Q26 86% → 1H26 89% → 공급부족 지속'),
          ('q27', '2027E Q 증가 (2026E 대비)', 0.05, F_PCT, '「27년부터 단계적 양산」이나 세종·베트남 완공은 2028년 4~5월 → 2027은 가동률·디보틀넥 위주'),
          ('p27', '2027E USD ASP 증가 (2H26E 대비)', 0.10, F_PCT, 'FCBGA 비중 확대(대면적 ㎡당 단가↑) + 공급부족 가격 인상 + CCL/PPG 원가(+17%) 전가. 2027 기판 LTA $267mn'),
          ('q28', '2028E 기존라인 Q 증가', 0.02, F_PCT, '기존 라인은 풀가동 유지'),
          ('p28', '2028E USD ASP 증가', 0.05, F_PCT, 'AI 기판 대형화 지속, 신규 Capa 유입으로 가격 상승폭 둔화'),
          ('new_ratio', '신규 Capa 규모 (2026E Capa 대비)', 0.50, F_PCT, '세종 4.27조 + 베트남 2.51조 = 6.78조 투자. 보수적으로 기존 면적 Capa의 50%로 가정'),
          ('ramp', '신규 Capa 2028년 평균 기여율', 0.35, F_PCT, '4~5월 완공 → 하반기 램프업(7개월 × 램프 수율) ≈ 35%'),
          ('fc_4q', 'FCBGA 비중 4Q26E', 0.63, F_PCT, '대신 3Q26 추정 62.5%, 신한 60~63%'),
          ('fc_27', 'FCBGA 비중 2027E', 0.67, F_PCT, 'AI·서버 FCBGA 성장 > BGA(모바일 AP·메모리)'),
          ('fc_28', 'FCBGA 비중 2028E', 0.72, F_PCT, '신규 Capa는 AI 서버용 FCBGA 중심')]
pq_extra = [dict(key='new', label='(참고) 2028E 신규 Capa 기여 면적 (천㎡)', fmt=F_NUM, vals={}),
            dict(key='total', label='패키지솔루션 매출 합계', style='bold', vals={})]
rows_extra = [dict(key='fc', label='FCBGA (서버CPU·AI가속기·네트워크·PC·전장)', indent=1, vals={}),
              dict(key='fcs', label='FCBGA 비중', style='pct', fmt=F_PCT, vals={}),
              dict(key='bga', label='BGA (모바일 AP·메모리·ARM CPU·SiP)', indent=1, vals={}),
              dict(key='fc_yoy', label='FCBGA YoY (%)', style='yoy', fmt=F_PCT, vals={}),
              dict(key='bga_yoy', label='BGA YoY (%)', style='yoy', fmt=F_PCT, vals={})]
pk_, PIN, PQB = seg_est_sheet('매출추정_패키지', 'PKG', '패키지솔루션(반도체 패키지기판) 매출 추정 — Q(천㎡) × P(USD/㎡) × 환율 + 2028 신규 Capa',
                              '천㎡', 1e3, 1, inputs, rows_extra, pq_extra, '패키지솔루션',
                              ['FCBGA 비중: 2024~3Q26은 대신증권 추정(R_REV), 4Q26E 이후 가정. 2023은 분해 자료 없음.'])
W = pk_.ws; R = lambda k: pk_.rows[('pq', k)]; Cc = lambda k: pk_.cols[('pq', k)]
# 2028E Q = 기존*(1+g) + 2026E capa * new_ratio * ramp
W[f"{Cc('2028E')}{R('new')}"] = f"={Cc('2026E')}{R('capa')}*{PIN('new_ratio')}*{PIN('ramp')}"
W[f"{Cc('2028E')}{R('q')}"] = f"={Cc('2027E')}{R('q')}*(1+{PIN('q28')})+{Cc('2028E')}{R('new')}"
for k, _ in PQ_COLS:
    W[f"{Cc(k)}{R('total')}"] = f"={Cc(k)}{R('krw')}"
rows = [dict(key='lta', label='공시된 기판 LTA 2027 귀속 ($mn)', fmt=F_USD, vals={'v': f"='R_공시'!${rg.cols[('ct','usd27')]}${rg.rows[('ct','pkg')]}"}),
        dict(key='usd27', label='당사 2027E 패키지 USD 매출 ($mn)', fmt=F_USD, vals={'v': f"={Cc('2027E')}{R('usd')}"}),
        dict(key='cov', label='공시 LTA 커버리지 (%)', style='pct', fmt=F_PCT, vals={}),
        dict(key='cons', label='증권사 평균 2027E 패키지 (원)', vals={'v': f"={cons('2027E_패키지', '매출추정_패키지')}"}),
        dict(key='cons_asp', label='→ 컨센이 함의하는 2027E ASP ($/㎡, 당사 Q·환율)', fmt=F_NUM2, vals={}),
        dict(key='cons_p', label='→ 컨센 함의 ASP 증가율 (2H26E 대비)', style='yoy', fmt=F_PCT, vals={})]
pk_.table('chk', '크로스체크 — LTA · 컨센 역산', '(체크)', [('v', '값')], rows)
V = pk_.cols[('chk', 'v')]; rcx = lambda k: pk_.rows[('chk', k)]
W[f"{V}{rcx('cov')}"] = f"={V}{rcx('lta')}/{V}{rcx('usd27')}"
W[f"{V}{rcx('cons_asp')}"] = f"={V}{rcx('cons')}/{fxa('2027E','매출추정_패키지')}/({Cc('2027E')}{R('q')}*1e3)"
W[f"{V}{rcx('cons_p')}"] = f"={V}{rcx('cons_asp')}/{Cc('2H26E')}{R('asp')}-1"
RA, CA = segA(pk_, 'PKG', 'total')
W = pk_.ws
fcmap = {'2024': '2024', '2025': '2025', '1Q26': '1Q26', '2Q26': '2Q26', '3Q26': '3Q26'}
for k, _ in EST_COLS:
    L = CA(k)
    if k == '2023':
        continue
    if k in fcmap:
        W[f"{L}{RA('fcs')}"] = f"={rr.ref('fc', 'share', fcmap[k], from_sheet='매출추정_패키지')}"
    elif k == '4Q26E': W[f"{L}{RA('fcs')}"] = f"={PIN('fc_4q')}"
    elif k == '2026E':
        W[f"{L}{RA('fcs')}"] = f"=SUMPRODUCT({CA('1Q26')}{RA('fcs')}:{CA('4Q26E')}{RA('fcs')},{CA('1Q26')}{RA('rev')}:{CA('4Q26E')}{RA('rev')})/{L}{RA('rev')}"
    elif k == '2027E': W[f"{L}{RA('fcs')}"] = f"={PIN('fc_27')}"
    elif k == '2028E': W[f"{L}{RA('fcs')}"] = f"={PIN('fc_28')}"
    W[f"{L}{RA('fc')}"] = f"={L}{RA('rev')}*{L}{RA('fcs')}"
    W[f"{L}{RA('bga')}"] = f"={L}{RA('rev')}-{L}{RA('fc')}"
for a, b in (('2025', '2024'), ('2026E', '2025'), ('2027E', '2026E'), ('2028E', '2027E')):
    for r_ in ('fc', 'bga'):
        W[f"{CA(a)}{RA(r_ + '_yoy')}"] = yoy_formula(f"{CA(a)}{RA(r_)}", f"{CA(b)}{RA(r_)}")
for q, pq in PREVQ.items():
    for r_, rk in (('fc', 'fc'), ('bga', 'bga')):
        W[f"{CA(q)}{RA(r_ + '_yoy')}"] = yoy_formula(f"{CA(q)}{RA(r_)}", rr.ref('fc', rk, pq, from_sheet='매출추정_패키지'))
pk_.text('부가설명 — 패키지솔루션', 'B', 10, mark='-')
for t in ['① 왜 이렇게 쪼갰나: 회사는 기판을 한 줄로 공시하지만 성장 동인이 정반대 — FCBGA(AI가속기·서버CPU, 공급부족·가격 상승) vs BGA(모바일 AP·메모리, 스마트폰 수요 연동). 증권사 추정(대신)으로 과거를 복원하고 비중 가정으로 미래를 나눔.',
          '② Q(천㎡): DART 생산능력 2025 810천㎡, 1H26 366천㎡(연환산 732) → MLCC처럼 대면적·고다층화로 「면적 Capa」가 오히려 감소(컨콜: Capa 잠식 심화). 가동률 2025 70% → 1H26 89%.',
          '③ P(USD/㎡): 2025 $2,860 → 1H26 $3,100 → 2H26E $4,040. 공시 ASP 1H26 +18.7% YoY, 원재료 CCL/PPG +17.4% → 가격 전가 + FCBGA 믹스.',
          '④ 2027E: 신규 공장(세종 4.27조·베트남 2.51조)은 2028년 4~5월 완공 → 2027 Q는 +5%만. 성장 대부분은 P(+10%).',
          '⑤ 2028E: 신규 Capa = 2026 면적 Capa의 50% × 2028 기여율 35% → 면적 +18%. 회사 「28년부터 매출 본격 기여」와 정합. 증설 규모·램프 속도가 상방 변수.',
          '⑥ 리스크: FCBGA는 소수 고객(빅테크) 의존 — 고객 내재화·경쟁사(Ibiden·Unimicron) 증설. 유리기판(글라셈 JV)은 2028년까지 매출 반영하지 않음.']:
    pk_.text(t, 'L', 9, indent=1)

# ================================================================ 광학
inputs = [('q4q', '4Q26E 생산(Q) QoQ', -0.08, F_PCT, '연말 재고조정 + 메모리 가격 급등에 따른 스마트폰 가격 인상 → 세트 수요 둔화(컨콜). 대신·미래·신한 4Q26 광학 −18% QoQ(원화) 수준'),
          ('q4p', '4Q26E USD ASP QoQ', -0.02, F_PCT, '전략거래선 차기 플래그십(고화소·Slim) 신제품 초기 물량은 1Q 집중 → 4Q 믹스 약화'),
          ('h2capa', '2H26 생산능력 증가 (1H26 대비)', 0.0, F_PCT, '광학 투자 1H26 391억(감가 수준) → Capa 유지'),
          ('h2util', '2H26 가동률', 0.66, F_PCT, '1H26 70%, 2025 66%. 하반기 수요 둔화'),
          ('q27', '2027E Q 증가 (2026E 대비)', -0.05, F_PCT, '스마트폰 출하 정체·부품 원가 상승(메모리) → 카메라 탑재 수량 감소. 전장·휴머노이드 성장분은 아직 작음'),
          ('p27', '2027E USD ASP 증가 (2026E 평균 대비)', 0.03, F_PCT, '2억화소 폴디드줌·Slim·전장 고화소 믹스 개선. ※광학은 계절성이 커 2H26이 아닌 2026E 평균을 기준으로 함'),
          ('q28', '2028E Q 증가', -0.02, F_PCT, '성숙 시장'),
          ('p28', '2028E USD ASP 증가', 0.03, F_PCT, '사양 상향 지속')]
pq_extra = [dict(key='total', label='광학솔루션 매출 합계', style='bold', vals={})]
rows_extra = [dict(key='sec', label='(참고) 삼성전자향 전사 매출 비중', style='pct', fmt=F_PCT, vals={})]
op_, OIN, OQB = seg_est_sheet('매출추정_광학', 'OPT', '광학솔루션(카메라모듈) 매출 추정 — Q(백만개) × P(USD/개) × 환율',
                              '백만개', 1e6, 1, inputs, rows_extra, pq_extra, '광학솔루션',
                              ['광학은 응용처(스마트폰/전장/로봇) 분해 수치가 공시·IR·리포트 어디에도 없어 단일 라인 유지 — 억지로 쪼개면 무근거 가정이 됨.'])
W = op_.ws; R = lambda k: op_.rows[('pq', k)]; Cc = lambda k: op_.cols[('pq', k)]
W[f"B{R('total') + 2}"] = 'P 증감: 2H26E는 1H26 대비, 2027E는 「2026E 평균」 대비(광학은 1Q 플래그십 효과로 계절성이 커서 연평균 기준), 2028E는 2027E 대비'
W[f"{Cc('2027E')}{R('asp')}"] = f"={Cc('2026E')}{R('asp')}*(1+{OIN('p27')})"
W[f"{Cc('2027E')}{R('asp_yoy')}"] = yoy_formula(f"{Cc('2027E')}{R('asp')}", f"{Cc('2026E')}{R('asp')}")
for k, _ in PQ_COLS:
    W[f"{Cc(k)}{R('total')}"] = f"={Cc(k)}{R('krw')}"
RA, CA = segA(op_, 'OPT', 'total')
W = op_.ws
for k, src in (('2023', '2023'), ('2024', '2024'), ('2025', '2025')):
    W[f"{CA(k)}{RA('sec')}"] = f"={rr.ref('sec', 's', src, from_sheet='매출추정_광학')}"
op_.text('부가설명 — 광학솔루션', 'B', 10, mark='-')
for t in ['① 구조: 카메라모듈. 주 고객 삼성전자(전사 매출의 26~32%가 삼성전자향, 상당 부분이 카메라로 추정) + 중화권 OEM + 글로벌 EV(전장) + 북미 휴머노이드(3Q26 양산 개시).',
          '② Q(백만개): 생산 2023 121 → 2025 98 → 1H26 52. 가동률 66~71%로 Capa 여유 → 매출은 「수요」가 결정.',
          '③ P(USD/개): 2023 $20.7 → 2024 $24.4 → 2025 $27.3 → 1H26 $27.3. 고화소·폴디드줌 믹스로 2년간 +32% 올랐지만 2026년 들어 정체(공시 ASP 원화 1H26 +2.4%).',
          '④ 2027E: 메모리 가격 급등 → 스마트폰 가격 인상 → 세트 수요 둔화(3Q26 컨콜). Q −5%, P +3% → 달러 매출 −2%, 환율 효과까지 원화 −9%.',
          '⑤ 증권사 2027E 광학 추정은 3.1~4.1조로 편차 최대 → 전장·로봇 성장 가정 차이. 당사는 근거 수치가 나오기 전까지 보수적으로 둠(상방 옵션).']:
    op_.text(t, 'L', 9, indent=1)

# ================================================================ REV (최종)
rv = S('삼성전기_REV')
wb.move_sheet('삼성전기_REV', offset=-(len(wb.sheetnames) - 1))
def seg_link(sh, rowk, k): return sh.ref('A', rowk, k, from_sheet='삼성전기_REV')
rows = [dict(key='rev', label='매출액', style='bold', vals={}),
        dict(key='yoy', label='YoY (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='comp', label='컴포넌트', style='bold', vals={}),
        dict(key='comp_s', label='% of sales', style='pct', fmt=F_PCT, vals={}),
        dict(key='comp_y', label='YoY (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='cap', label='Capacitors(MLCC)', indent=2, vals={}),
        dict(key='ele', label='전자소자', indent=2, vals={}),
        dict(key='si', label='실리콘 커패시터(신규)', indent=2, vals={}),
        dict(key='pkg', label='패키지솔루션', style='bold', vals={}),
        dict(key='pkg_s', label='% of sales', style='pct', fmt=F_PCT, vals={}),
        dict(key='pkg_y', label='YoY (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='fc', label='FCBGA', indent=2, vals={}),
        dict(key='bga', label='BGA', indent=2, vals={}),
        dict(key='opt', label='광학솔루션', style='bold', vals={}),
        dict(key='opt_s', label='% of sales', style='pct', fmt=F_PCT, vals={}),
        dict(key='opt_y', label='YoY (%)', style='yoy', fmt=F_PCT, vals={})]
rv.table('rev', '최종 매출 추정 Table', '(단위: 백만 원)', EST_COLS, rows, final=True, est_from='4Q26E', vlines=['1Q26', '4Q26E', '2026E'],
         note=['실적: DART(2023~2Q26), 3Q26 부문은 IR(억원). 2026E = 1~3Q 실적 + 4Q26E. Capacitors/전자소자는 IR 비중, FCBGA/BGA는 증권사 추정 비중으로 분해.'])
W = rv.ws; RR = lambda k: rv.rows[('rev', k)]; CR = lambda k: rv.cols[('rev', k)]
for k, _ in EST_COLS:
    L = CR(k)
    W[f"{L}{RR('comp')}"] = f"={seg_link(cp_, 'rev', k)}"
    W[f"{L}{RR('cap')}"] = f"={seg_link(cp_, 'cap', k)}"
    W[f"{L}{RR('ele')}"] = f"={seg_link(cp_, 'ele', k)}"
    W[f"{L}{RR('si')}"] = f"={seg_link(cp_, 'si', k)}"
    W[f"{L}{RR('pkg')}"] = f"={seg_link(pk_, 'rev', k)}"
    if k != '2023':
        W[f"{L}{RR('fc')}"] = f"={seg_link(pk_, 'fc', k)}"
        W[f"{L}{RR('bga')}"] = f"={seg_link(pk_, 'bga', k)}"
    W[f"{L}{RR('opt')}"] = f"={seg_link(op_, 'rev', k)}"
    W[f"{L}{RR('rev')}"] = f"={L}{RR('comp')}+{L}{RR('pkg')}+{L}{RR('opt')}"
    for s_ in ('comp', 'pkg', 'opt'):
        W[f"{L}{RR(s_ + '_s')}"] = f"={L}{RR(s_)}/{L}{RR('rev')}"
    W[f"{L}{RR('comp_y')}"] = f"={seg_link(cp_, 'yoy', k)}"
    W[f"{L}{RR('pkg_y')}"] = f"={seg_link(pk_, 'yoy', k)}"
    W[f"{L}{RR('opt_y')}"] = f"={seg_link(op_, 'yoy', k)}"
prev_tot = {'2023': aref('TOTAL', '2022', '삼성전기_REV'), '2024': f"{CR('2023')}{RR('rev')}", '2025': f"{CR('2024')}{RR('rev')}",
            '1Q26': qref('TOTAL', '1Q25', '삼성전기_REV'), '2Q26': qref('TOTAL', '2Q25', '삼성전기_REV'), '3Q26': qref('TOTAL', '3Q25', '삼성전기_REV'),
            '4Q26E': qref('TOTAL', '4Q25', '삼성전기_REV'), '2026E': f"{CR('2025')}{RR('rev')}", '2027E': f"{CR('2026E')}{RR('rev')}", '2028E': f"{CR('2027E')}{RR('rev')}"}
for k, p in prev_tot.items():
    W[f"{CR(k)}{RR('yoy')}"] = yoy_formula(f"{CR(k)}{RR('rev')}", p)
# 컨센 비교
ccols = [('4Q26E', '4Q26E'), ('2026E', '2026E'), ('2027E', '2027E'), ('2028E', '2028E')]
rows = [dict(key='ours', label='당사 추정 매출액', style='bold', fmt=F_WON_B, vals={}),
        dict(key='fn', label='FnGuide 컨센서스', fmt=F_WON_B, vals={}),
        dict(key='fn_d', label='차이 (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='br', label='증권사 평균 (10/29 리뷰 5개사)', fmt=F_WON_B, vals={}),
        dict(key='br_d', label='차이 (%)', style='yoy', fmt=F_PCT, vals={}),
        dict(key='c_comp', label='컴포넌트 차이 vs 증권사 평균 (%)', style='pct', fmt=F_PCT, vals={}),
        dict(key='c_pkg', label='패키지 차이 vs 증권사 평균 (%)', style='pct', fmt=F_PCT, vals={}),
        dict(key='c_opt', label='광학 차이 vs 증권사 평균 (%)', style='pct', fmt=F_PCT, vals={})]
rv.table('cmp', '컨센서스 대비', '(단위: 십억 원)', ccols, rows, est_from='4Q26E',
         note='FnGuide: 2026E는 11/10 조회치, 2027E는 10/28 기준(리포트 인용), 4Q26E·2028E는 미확보. 증권사 평균은 컨센 시트 참조.')
W = rv.ws; RC = lambda k: rv.rows[('cmp', k)]; CC = lambda k: rv.cols[('cmp', k)]
for k, _ in ccols:
    L = CC(k)
    W[f"{L}{RC('ours')}"] = f"={CR(k)}{RR('rev')}"
    if k == '2026E': W[f"{L}{RC('fn')}"] = f"={cs.ref('fn','rev','2026E',absolute=True,from_sheet='삼성전기_REV')}"
    if k == '2027E': W[f"{L}{RC('fn')}"] = f"={cs.ref('fn2','rev','2027E',absolute=True,from_sheet='삼성전기_REV')}"
    W[f"{L}{RC('fn_d')}"] = f"=IFERROR({L}{RC('ours')}/{L}{RC('fn')}-1,\"\")"
    W[f"{L}{RC('br')}"] = f"={cons(k + '_합계', '삼성전기_REV')}"
    W[f"{L}{RC('br_d')}"] = f"=IFERROR({L}{RC('ours')}/{L}{RC('br')}-1,\"\")"
    for s_, nm in (('comp', '컴포넌트'), ('pkg', '패키지'), ('opt', '광학')):
        W[f"{L}{RC('c_' + s_)}"] = f"=IFERROR({CR(k)}{RR(s_)}/{cons(k + '_' + nm, '삼성전기_REV')}-1,\"\")"
# 환율 민감도
scen = [1250, 1300, 1340, 1400, 1450, 1500]
rows = [dict(key=f"fx{v}", label=f"USD/KRW {v:,}원", fmt=F_WON_B, vals={}) for v in scen]
rv.table('sens', '환율 민감도 — 2027E·2028E 매출 (달러 매출 고정, 환율만 변경)', '(단위: 십억 원)', [('2027E', '2027E'), ('2028E', '2028E')], rows,
         note='부문 매출 전액을 달러 연동으로 단순화(내수 3~4%). 기준 가정 1,340원. 환율 ±100원 ≈ 매출 ±7.5%')
for v in scen:
    for k in ('2027E', '2028E'):
        rv.ws[f"{rv.cols[('sens', k)]}{rv.rows[('sens', f'fx{v}')]}"] = f"={CR(k)}{RR('rev')}*{v}/{fxa(k, '삼성전기_REV')}"
rv.text('부가설명', 'B', 10, mark='-')
for t in ['① 결론: 2026E 14.3조(컨센 −1%), 2027E 17.3조(증권사 평균 −9%), 2028E 19.7조(평균 −18%). 방향(고성장)은 컨센과 같고 「속도」가 다름.',
          '② 컨센과의 갭 분해: (a) 환율 — 1,340원 고정(증권사는 대부분 1,400원대 가정 추정) (b) MLCC ASP — 당사 2027 +12%(2H26 대비) vs 컨센 함의 +20%대 (c) 기판 신규 Capa 램프 — 2028 기여율 35%로 보수적.',
          '③ 매출 구조(2026E): 컴포넌트 49% / 광학 27% / 패키지 24% → 2028E 컴포넌트 55% / 패키지 27% / 광학 18%. AI(MLCC·FCBGA)가 성장 전부, 광학은 역성장.',
          '④ 업데이트 포인트: 3Q26 분기보고서(11/16까지) 부문·Capa 확정치, 4Q 추가 LTA 공시, 무라타 실적·가격 정책, 환율.']:
    rv.text(t, 'L', 9, indent=1)


# ================================================================ Breakdown_논리
bd = S('Breakdown_논리', widths={'B': 26, 'data': 11})
bd.text('매출 Breakdown & 라인별 추정 논리 (과제 2·3번)', 'H', 12, RED, mark='X')
bd.text('원칙: DART 3개 부문을 「서로 다른 드라이버로 움직이는, 숫자로 확인 가능한」 크기까지만 쪼갠다. 근거 데이터가 없는 분해(예: 광학의 스마트폰/전장)는 하지 않는다.', 'L', 9, GRAY_M)
bd.gap()
hdr = ['라인', '2025A', '2026E', '2027E', '2028E', '쪼갠 근거 (데이터)', 'Q 드라이버', 'P 드라이버', '핵심 가정 (2027E)', '크로스체크 / 리스크']
r0 = bd.r
for j, h in enumerate(hdr):
    c = bd.ws.cell(row=r0, column=2 + j, value=h); c.font = fnt('B', 9, 'FFFFFF'); c.fill = PatternFill('solid', fgColor=RED)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = Border(bottom=dbl)
REVR = lambda key: rv.rows[('rev', key)]
lines = [
 ('컴포넌트', 'comp', '', '', '', '', '', True),
 ('  Capacitors(MLCC)', 'cap', 'IR 분기별 Capacitor 비중(89~92%) × DART 컴포넌트 매출',
  'DART 생산능력(억개)×가동률. 고적층 믹스로 개수 Capa는 감소 추세 → 필리핀 신공장(27)·부산',
  '매출÷생산량 = USD/천개. AI·전장 믹스 + LTA 가격 + 범용품 가격 인상', 'Q +6%, P +12%(2H26 평균 대비 = 4Q26 출구 대비 +7%)',
  '공시 MLCC LTA $2.35bn = 당사 27E의 38%(회사: LTA 과반) / 컨센은 P +29% 함의 / 리스크: 환율, IT 수요', False),
 ('  전자소자', 'ele', 'IR 비중의 나머지(8~10%)', 'MLCC와 같은 라인·고객 → 동일 Q×P 안에서 비중 고정', '—', '비중 8% 유지', '규모 작아 별도 모델 실익 없음', False),
 ('  실리콘 커패시터', 'si', '단일판매·공급계약 공시(2026-05-20, $1,035.4mn, 27~28)', '계약 물량', '계약 금액', '$517.7mn/년 × 1,340원', '계약 외 추가 수주는 상방, 기술·수율 리스크', False),
 ('패키지솔루션', 'pkg', '', '', '', '', '', True),
 ('  FCBGA', 'fc', '증권사 추정(대신) 분기 FCBGA/BGA → 비중', 'DART 면적 Capa(천㎡)×가동률 + 2028 세종·베트남 신규 Capa(6.78조)',
  'USD/㎡: 대면적·고다층 AI 기판 믹스 + 공급부족 + CCL/PPG 원가 전가', '비중 67%, 패키지 Q +5% / P +10%; 2028 신규 Capa 50%×기여율 35%',
  '공시 기판 LTA $267mn / 컨센은 P +23% 함의 / 리스크: 고객 집중, 경쟁사 증설', False),
 ('  BGA', 'bga', '패키지 − FCBGA', '스마트폰 AP·메모리·ARM CPU 수요', '원가(CCL/PPG +17%) 전가', '잔여 = 패키지×(1−FCBGA 비중) → 연 +4~7%', '스마트폰 출하 둔화', False),
 ('광학솔루션', 'opt', '분해 근거 없음 → 단일 라인', 'DART 생산량(백만개). 가동률 66~70%로 Capa 여유 → 수요가 결정',
  'USD/개: 고화소·폴디드줌·Slim, 전장 고화소', 'Q −5%, P +3% (2026E 평균 대비)', '증권사 편차 최대(3.1~4.1조) / 상방: 전장·휴머노이드', True),
 ('합계', 'rev', '', '', '', '공통 환율 1,340원 고정(수출 97%)', '컨센 대비 26E −1% / 27E −9% / 28E −18%', True)]
r = r0 + 1
for lab, key, why, qd, pd_, asm, chk, bold in lines:
    lc = bd.ws.cell(row=r, column=2, value=lab); lc.font = fnt('B' if bold else 'L', 9); lc.alignment = Alignment(vertical='top')
    for j, k in enumerate(['2025', '2026E', '2027E', '2028E']):
        c = bd.ws.cell(row=r, column=3 + j, value=f"='삼성전기_REV'!{rv.cols[('rev', k)]}{REVR(key)}")
        c.number_format = F_WON_B; c.font = Font(name='Arial', size=9, bold=bold); c.alignment = Alignment(vertical='top')
    for j, t in enumerate([why, qd, pd_, asm, chk]):
        c = bd.ws.cell(row=r, column=7 + j, value=t); c.font = fnt('L', 8); c.alignment = Alignment(wrap_text=True, vertical='top')
    bd.ws.row_dimensions[r].height = 15 if bold and not why else 48
    for j in range(2, 12):
        bd.ws.cell(row=r, column=j).border = Border(bottom=hair if not bold else thin)
    r += 1
for j in range(2, 12): 
    c = bd.ws.cell(row=r - 1, column=j); c.border = Border(bottom=thick)
bd.ws.cell(row=r, column=2, value='(단위: 십억 원) 금액은 삼성전기_REV 연동').font = fnt('L', 8, GRAY_M)
for col, w in zip('GHIJK', (30, 34, 30, 30, 34)): bd.ws.column_dimensions[col].width = w
bd.r = r + 2
bd.text('왜 이 크기인가', 'B', 10, mark='-')
for t in ['① MLCC vs 전자소자: 회사가 IR에서 비중을 매 분기 공개 → 「근거 있는 분해」. MLCC만이 AI·LTA·가격 사이클의 본체.',
          '② 실리콘 커패시터: 계약 공시로 매출이 숫자로 확정되는 유일한 신규 품목 → 계약금액 그대로 반영(추정 아닌 「확인」).',
          '③ FCBGA vs BGA: 같은 기판이지만 고객(빅테크 AI vs 스마트폰·메모리)과 사이클이 정반대. 회사 미공시라 증권사 추정치(대신) 사용을 명시.',
          '④ 광학은 쪼개지 않음: 스마트폰/전장/로봇 비중 수치가 공시·IR·리포트 어디에도 없음 → 억지 분해는 「무근거 가정」(매출교육 C7).',
          '⑤ 더 쪼갤 후보(다음 단계): MLCC 응용처(AI·데이터센터 / 전장 / IT) — 현재 근거는 「3Q26 컴포넌트 내 데이터센터 30%(IBK)」 한 점뿐 → 전방(엔비디아·ASIC 출하, CoWoS Capa) 데이터 확보 시 분해.']:
    bd.text(t, 'L', 9, indent=1)

# 시트 순서: REV, 매출추정 3, 컨센, FX, >>>Raw, R_*
order = ['삼성전기_REV', 'Breakdown_논리', '매출추정_컴포넌트', '매출추정_패키지', '매출추정_광학', '컨센', 'FX', 'R_IS', 'R_REV', 'R_CAPA', 'R_공시']
raw_div = wb.create_sheet('>>>Raw')
raw_div['B2'] = '아래는 DART·IR·공시 원자료 시트 (R_)'; raw_div['B2'].font = fnt('B', 10)
order.insert(7, '>>>Raw')
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 85
    ws.freeze_panes = 'C1' if ws.title not in ('>>>Raw',) else None
wb.active = 0
wb.save(OUT)
print('saved', OUT)
