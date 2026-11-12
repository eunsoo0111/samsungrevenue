"""삼성전기 매출 시트 3장(삼성전기_RAW / 삼성전기_REV / 삼성전기_매출추정)을 다시 만든다.

사용법:  python3 scripts/build_samsung_sheets.py 입력.xlsx 출력.xlsx

- RAW는 DART 정기보고서 표(사업의 내용 II-2·3·4, 연결주석 3·20-1)를 표 모양 그대로 옮긴 것이다.
- 금액은 원 단위로 저장하고 표시형식(#,##0,,)으로 백만 원을 보여준다.
- REV·매출추정의 실적 칸은 전부 삼성전기_RAW 셀 하나를 그대로 참조한다.
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color

RAW, REV, EST = "삼성전기_RAW", "삼성전기_REV", "삼성전기_매출추정"

# ── RAW 기간 열 (DART 표를 옆으로 이어 붙인 순서). 2분기 열만 계산값(반기 − 1분기).
PERIOD_COLS = [
    ("F", "제50기", "('22.01.01~'22.12.31)"), ("G", "제51기", "('23.01.01~'23.12.31)"),
    ("H", "제52기", "('24.01.01~'24.12.31)"), ("I", "제53기", "('25.01.01~'25.12.31)"),
    ("J", "제54기 1분기", "('26.01.01~'26.03.31)"), ("K", "제54기 반기", "('26.01.01~'26.06.30)"),
    ("L", "제54기 2분기", "(반기 − 1분기)"),
    ("M", "제53기 1분기", "('25.01.01~'25.03.31)"), ("N", "제53기 반기", "('25.01.01~'25.06.30)"),
    ("O", "제53기 2분기", "(반기 − 1분기)"),
]
# REV·매출추정이 참조하는 열
COL = {"2022": "F", "2023": "G", "2024": "H", "2025": "I", "1Q26": "J", "2Q26": "L", "1Q25": "M", "2Q25": "O"}
PRIOR = {"2023": "2022", "2024": "2023", "2025": "2024", "1Q26": "1Q25", "2Q26": "2Q25"}
N = None  # 미공시

# ── DART 숫자 (금액 단위: 백만 원, 공시 그대로). (a, b)는 반기 누적 a − 1분기 b.
# 1. 사업부문별 매출실적 — II. 사업의 내용 > 4. 매출 및 수주상황 > 가. 매출실적
SALES = {
    "comp_exp": [3983498, 3762448, 4307441, 5006765, 1350693, (2930083, 1350693), 1173562, (2408692, 1173562)],
    "comp_dom": [148785, 140574, 154609, 191737, 57805, (127836, 57805), 42704, (88290, 42704)],
    "comp": [4132283, 3903022, 4462050, 5198502, 1408498, (3057919, 1408498), 1216266, (2496982, 1216266)],
    "pkg_exp": [2051622, 1688535, 2006869, 2266771, 701266, (1447667, 701266), 492534, (1049027, 492534)],
    "pkg_dom": [36697, 28843, 27788, 34982, 23752, (48982, 23752), 6868, (14948, 6868)],
    "pkg": [2088319, 1717378, 2034657, 2301753, 725018, (1496649, 725018), 499402, (1063975, 499402)],
    "opt_exp": [2997380, 2993994, 3534086, 3615414, 1060332, (2087151, 1060332), 956020, (1852869, 956020)],
    "opt_dom": [188628, 278018, 263310, 198790, 15265, (24599, 15265), 66961, (109452, 66961)],
    "opt": [3186008, 3272012, 3797396, 3814204, 1075597, (2111750, 1075597), 1022981, (1962321, 1022981)],
    "tot_exp": [9032500, 8444977, 9848396, 10888950, 3112291, (6464901, 3112291), 2622116, (5310588, 2622116)],
    "tot_dom": [374110, 447435, 445707, 425509, 96822, (201417, 96822), 116533, (212690, 116533)],
    "tot": [9406610, 8892412, 10294103, 11314459, 3209113, (6666318, 3209113), 2738649, (5523278, 2738649)],
}
# 2. 매출 유형별 매출실적 — II-4 (합계) / 3. 부문별 수익 구분 — 연결주석 20-1
TYPE = {
    "comp_goods": [N, 3905722, 4443098, 5194938, 1405602, (3049394, 1405602), 1199835, (2472406, 1199835)],
    "comp_other": [N, -2700, 18952, 3563, 2896, (8525, 2896), 16431, (24576, 16431)],
    "comp_sum": [N, 3903022, 4462050, 5198501, 1408498, (3057919, 1408498), 1216266, (2496982, 1216266)],
    "pkg_goods": [N, 1460016, 1680718, 1999104, 637832, (1337404, 637832), 426565, (913374, 426565)],
    "pkg_other": [N, 257362, 353939, 302649, 87186, (159245, 87186), 72837, (150601, 72837)],
    "pkg_sum": [N, 1717378, 2034657, 2301753, 725018, (1496649, 725018), 499402, (1063975, 499402)],
    "opt_goods": [N, 3252588, 3747133, 3785921, 1070374, (2098802, 1070374), 1015179, (1941904, 1015179)],
    "opt_other": [N, 19424, 50263, 28284, 5223, (12948, 5223), 7802, (20417, 7802)],
    "opt_sum": [N, 3272012, 3797396, 3814205, 1075597, (2111750, 1075597), 1022981, (1962321, 1022981)],
    "tot_goods": [9255753, 8618326, 9870949, 10979963, 3113808, (6485600, 3113808), 2641579, (5327684, 2641579)],
    "tot_other": [150857, 274086, 423154, 334496, 95305, (180718, 95305), 97070, (195594, 97070)],
}
NOTE_TOTAL = [N, 8892412, 10294103, 11314459, 3209113, (6666318, 3209113), 2738649, (5523278, 2738649)]
# 3. 지역별 매출 — 연결주석 3. 영업부문정보 > 지역에 대한 공시 (전사 기준)
REGION = {
    "kr_dom": [N, 447435, 445707, 425509, 96823, (201427, 96823), 116530, (212691, 116530)],
    "kr_exp": [N, 2369503, 3022024, 3201268, 825487, (1746875, 825487), 695841, (1544106, 695841)],
    "cn": [N, 3393488, 4054406, 4441764, 1176228, (2554907, 1176228), 1058231, (2104718, 1058231)],
    "sea": [N, 1917253, 1990016, 2368787, 864750, (1620739, 864750), 673821, (1238620, 673821)],
    "am": [N, 368331, 363763, 422280, 124982, (266761, 124982), 96405, (198014, 96405)],
    "eu": [N, 295063, 328321, 371139, 100893, (213122, 100893), 81750, (179196, 81750)],
    "jp": [N, 101339, 89866, 83712, 19950, (62487, 19950), 16071, (45933, 16071)],
}
# 4. 주요 매출처 — 연결주석 3 > 주요 고객에 대한 공시 (억원 단위 공시)
SEC = [N, 2873200, 2985500, 3093800, 949600, (1710800, 949600), 921200, (1640400, 921200)]
# 5. 평균판매가격 변동률(전년 대비) — II-2 나. 가격변동추이. 2Q 열은 반기 누적 기준 공시값.
ASP = {
    "comp": [0.302, -0.196, 0.042, 0.009, 0.081, 0.139, -0.005, 0.003],
    "pkg": [0.259, 0.036, 0.081, 0.041, 0.144, 0.187, -0.009, 0.0],
    "opt": [-0.041, -0.043, 0.096, 0.140, 0.023, 0.024, 0.212, 0.173],
}
# 6. 생산능력·생산실적 — II-3 다·라 (공시 단위 그대로: 억개 / 천㎡ / 백만개)
PROD = {
    "comp_cap": [12445, 12568, 12419, 11159, 2953, (6156, 2953), 2606, (5317, 2606)],
    "comp_out": [7227, 8809, 10018, 10363, 2805, (5614, 2805), 2503, (5189, 2503)],
    "pkg_cap": [784, 784, 788, 810, 183, (366, 183), 197, (399, 197)],
    "pkg_out": [699, 458, 509, 565, 158, (325, 158), 125, (272, 125)],
    "opt_cap": [195, 158, 158, 149, 37, (74, 37), 37, (74, 37)],
    "opt_out": [115, 121, 114, 98, 26, (52, 26), 22, (46, 22)],
}

# 가동률 — II-3 라 (2). 공시값 그대로 (분기 열은 1분기 / 반기 누적).
UTIL = {
    "comp": [0.58, 0.70, 0.81, 0.93, 0.95, 0.91, 0.96, 0.98],
    "pkg": [0.89, 0.58, 0.65, 0.70, 0.86, 0.89, 0.63, 0.68],
    "opt": [0.59, 0.76, 0.72, 0.66, 0.71, 0.70, 0.60, 0.62],
}

# ── 서식 (닌텐도 파일과 동일한 폰트·표시형식)
MN = r"#,##0,,;\(#,##0,,\);\-"
MN1 = r"#,##0.0,,;\(#,##0.0,,\);\-"
PCT = "0.0%"
QTY = "#,##0"
GREY = "FF808080"
THIN, DBL, MED = Side(style="thin"), Side(style="double"), Side(style="medium")


def font(name="Arial", size=11, bold=False, italic=False, theme=None, tint=0.0, rgb=None):
    color = Color(rgb=rgb) if rgb else (Color(theme=theme, tint=tint) if theme is not None else None)
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def kfont(weight="M", theme=1):
    return font(name=f"AppleSDGothicNeo{weight}00", theme=theme)


def fill(theme):
    return PatternFill(fill_type="solid", fgColor=Color(theme=theme))


def put(ws, ref, value=None, fnt=None, nf=None, align=None, indent=0, fil=None, **border):
    c = ws[ref]
    if value is not None:
        c.value = value
    if fnt is not None:
        c.font = fnt
    if nf is not None:
        c.number_format = nf
    if align or indent:
        c.alignment = Alignment(horizontal=align, indent=indent)
    if fil is not None:
        c.fill = fil
    if border:
        c.border = Border(**border)
    return c


def fresh_sheet(wb, title):
    """같은 자리·같은 탭 색으로 빈 시트를 다시 만든다."""
    old = wb[title]
    idx, tab = wb.index(old), old.sheet_properties.tabColor
    wb.remove(old)
    ws = wb.create_sheet(title, idx)
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 80
    return ws


# ───────────────────────────── 삼성전기_RAW ─────────────────────────────
# v1처럼 DART 표 모양 그대로 (굴림, 회색 머리글, 실선 테두리). 기간 열만 옆으로 이어 붙인다.
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill(fill_type="solid", fgColor="FFD7D7D7")
ZERO_PCT = r"0.0%;\-0.0%;0.0%"


def dfont(bold=False, size=11):
    return Font(name="굴림", size=size, bold=bold, color="FF000000")


def build_raw(wb):
    ws = fresh_sheet(wb, RAW)
    ws.sheet_view.zoomScale = 85
    for col, width in {"A": 2.6, "B": 18, "C": 14, "D": 20, "E": 12}.items():
        ws.column_dimensions[col].width = width
    for col, *_ in PERIOD_COLS:
        ws.column_dimensions[col].width = 19

    rows = {}  # key -> 행 번호 (REV·매출추정이 참조)
    state = {"r": 6}

    def box(rng, value=None, align="center", fil=None, nf=None):
        """범위를 병합하고 테두리를 두른 칸 하나."""
        cells = [c for row in ws[rng] for c in row] if ":" in rng else [ws[rng]]
        for c in cells:
            c.border, c.font = BOX, dfont()
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
            if fil:
                c.fill = fil
        if len(cells) > 1:
            ws.merge_cells(rng)
        if value is not None:
            cells[0].value = value
        if nf:
            cells[0].number_format = nf

    def table(title, unit, labels):
        """표 제목 + 머리글 2행. labels: [(시작열, 끝열, 글자)]"""
        r = state["r"]
        ws[f"B{r}"].value, ws[f"B{r}"].font = title, dfont(bold=True)
        ws[f"O{r}"].value, ws[f"O{r}"].font = unit, dfont()
        ws[f"O{r}"].alignment = Alignment(horizontal="right")
        for a, b, text in labels:
            box(f"{a}{r + 1}:{b}{r + 2}", text, fil=HEAD_FILL)
        for col, name, span in PERIOD_COLS:
            box(f"{col}{r + 1}", name, fil=HEAD_FILL)
            box(f"{col}{r + 2}", span, fil=HEAD_FILL)
        state["r"] = r + 3

    def values(key, vals, nf=MN, scale=10**6, second="diff"):
        """기간 값 한 줄. vals = [제50기, 51기, 52기, 53기, 1Q26, (반기26, 1Q26) 또는 반기26, 1Q25, (반기25, 1Q25) 또는 반기25]"""
        r = state["r"]
        rows[key] = r
        half = lambda v: v[0] if isinstance(v, tuple) else v
        cells = dict(zip("FGHIJKMN", [*vals[:5], half(vals[5]), vals[6], half(vals[7])]))
        for col, v in cells.items():
            box(f"{col}{r}", None if v is None else v * scale, align="right", nf=nf)
        for col, a, b in (("L", "K", "J"), ("O", "N", "M")):
            if second == "diff":
                f = f"={a}{r}-{b}{r}"  # 2분기 = 반기 − 1분기
            else:
                f = second(col) if second else None
            box(f"{col}{r}", f, align="right", nf=nf)
        state["r"] = r + 1
        return r

    def end_table():
        state["r"] += 2

    put(ws, "B2", "삼성전기 DART 공시 표 (연결 기준)", dfont(bold=True, size=12))
    put(ws, "B3", "- 표 모양과 항목은 DART 그대로, 여러 보고서의 기간 열만 옆으로 이어 붙임. 2분기 열만 계산값 (반기 − 1분기).", dfont())
    put(ws, "B4", "- 금액은 원 단위로 입력하고 표시만 백만원 (DART 숫자 × 1,000,000). 제50·51기는 중단사업 반영 재작성 수치.", dfont())
    put(ws, "B5", "- 제50기의 3~5번 표(주석)는 재작성 기준 미공시라 공란.", dfont())
    state["r"] = 7

    seg = [("comp", "컴포넌트", "수동소자\n(MLCC, Inductor,\nChip Resistor 등)"),
           ("pkg", "패키지솔루션", "반도체패키지기판"),
           ("opt", "광학솔루션", "카메라모듈")]
    won = "(단위 : 백만원)"

    # 1. 매출실적
    table("1. 매출실적 — 사업보고서 II. 사업의 내용 > 4. 매출 및 수주상황 > 가. 매출실적", won,
          [("B", "B", "사업부문"), ("C", "C", "매출유형"), ("D", "D", "품 목"), ("E", "E", "구 분")])
    for k, name, product in seg + [("tot", "합 계", None)]:
        r = state["r"]
        if product:
            box(f"B{r}:B{r + 2}", name)
            box(f"C{r}:C{r + 2}", "제품, 상품,\n용역, 기타매출")
            box(f"D{r}:D{r + 2}", product)
        else:
            box(f"B{r}:D{r + 2}", name)
        for suffix, label in (("_exp", "수출"), ("_dom", "내수"), ("", "합계")):
            box(f"E{state['r']}", label)
            values(k + suffix, SALES[k + suffix])
    end_table()

    # 2. 매출 유형별 매출실적
    table("2. 매출 유형별 매출실적 — 사업보고서 II-4 가. 매출실적", won, [("B", "E", "구 분")])
    for key, label, vals in (("tot_goods", "제품/상품", TYPE["tot_goods"]), ("tot_other", "용역 및 기타매출", TYPE["tot_other"]),
                             ("type_sum", "계", SALES["tot"])):
        box(f"B{state['r']}:E{state['r']}", label)
        values(key, vals)
    end_table()

    # 3. 부문별 수익 구분
    table("3. 고객과의 계약에서 생기는 수익의 구분 (부문별) — 연결재무제표 주석 20-1", won,
          [("B", "B", "사업부문"), ("C", "E", "구 분")])
    for k, name, _ in seg + [("tot", "부문 합계", None)]:
        r = state["r"]
        box(f"B{r}:B{r + 2}", name)
        for key, label, vals in ((f"{k}_goods", "제품 및 상품 매출", TYPE[f"{k}_goods"]),
                                 (f"{k}_other", "기타", TYPE[f"{k}_other"]),
                                 (f"{k}_sum", "합계", NOTE_TOTAL if k == "tot" else TYPE[f"{k}_sum"])):
            box(f"C{state['r']}:E{state['r']}", label)
            values(f"note_{key}" if k == "tot" else key, [N, *vals[1:]] if k == "tot" else vals)
    end_table()

    # 4. 지역별 수익
    table("4. 지역에 대한 공시 (지역별 수익, 전사 기준) — 연결재무제표 주석 3. 영업부문정보", won,
          [("B", "B", "지 역"), ("C", "E", "구 분")])
    r = state["r"]
    box(f"B{r}:B{r + 1}", "국내")
    box(f"B{r + 2}:B{r + 6}", "외국")
    for key, label in (("kr_dom", "내수"), ("kr_exp", "수출"), ("cn", "중국"), ("sea", "동남아"), ("am", "미주"),
                       ("eu", "유럽"), ("jp", "일본")):
        box(f"C{state['r']}:E{state['r']}", label)
        values(key, REGION[key])
    box(f"B{state['r']}:E{state['r']}", "지역 합계")
    values("region_sum", NOTE_TOTAL)
    end_table()

    # 5. 주요 고객
    table("5. 주요 고객에 대한 공시 — 연결재무제표 주석 3. 영업부문정보 (억원 단위 공시)", won, [("B", "E", "구 분")])
    box(f"B{state['r']}:E{state['r']}", "삼성전자와 그 종속기업 수익")
    values("sec", SEC)
    end_table()

    # 6. 가격변동추이 (DART는 문장으로 공시)
    table("6. 주요 제품 등의 가격변동추이 (평균판매가격, 전년 대비) — 사업보고서 II-2 나. (문장 공시, 반기는 누적 기준)",
          "(단위 : %)", [("B", "B", "사업부문"), ("C", "E", "주요 제품")])
    for k, name, product in (("comp", "컴포넌트", "MLCC"), ("pkg", "패키지솔루션", "반도체패키지기판"),
                             ("opt", "광학솔루션", "카메라모듈")):
        box(f"B{state['r']}", name)
        box(f"C{state['r']}:E{state['r']}", product)
        values(f"asp_{k}", ASP[k], nf=ZERO_PCT, scale=1, second=None)
    end_table()

    # 7~9. 생산능력 / 생산실적 / 가동률
    units = (("comp", "컴포넌트\n(억개)", "수동소자\n(MLCC, Inductor, Chip Resistor 등)"),
             ("pkg", "패키지솔루션\n(천㎡)", "반도체패키지기판"), ("opt", "광학솔루션\n(백만개)", "카메라모듈"))
    for n, title, suffix in ((7, "생산능력 — 사업보고서 II-3 다. 생산능력 및 산출근거 (글로벌 완제품 기준)", "_cap"),
                             (8, "생산실적 — 사업보고서 II-3 라. (1) 생산실적 (글로벌 완제품 기준)", "_out"),
                             (9, "평균가동률 — 사업보고서 II-3 라. (2) 가동률 (생산능력 대비 생산실적, 2분기 열은 8번 ÷ 7번)", "_util")):
        table(f"{n}. {title}", "(단위 : 사업부문별 표기)" if n < 9 else "(단위 : %)",
              [("B", "B", "사업부문"), ("C", "D", "품 목"), ("E", "E", "사업소")])
        box(f"E{state['r']}:E{state['r'] + 2}", "국내 및\n해외사업장")
        for k, name, product in units:
            r = state["r"]
            box(f"B{r}", name)
            box(f"C{r}:D{r}", product)
            ws.row_dimensions[r].height = 30
            if n < 9:
                values(k + suffix, PROD[k + suffix], nf=QTY, scale=1)
            else:
                values(k + suffix, UTIL[k], nf="0%", scale=1,
                       second=lambda col, k=k: f"={col}{rows[k + '_out']}/{col}{rows[k + '_cap']}")
        end_table()
    return rows

# ───────────────────── 삼성전기_REV / 삼성전기_매출추정 공통 ─────────────────────
YEARS = [("C", 2023, "2023"), ("D", 2024, "2024"), ("E", 2025, "2025"), ("F", "1Q26", "1Q26"),
         ("G", "2Q26", "2Q26"), ("H", 2026, None), ("I", 2027, None), ("J", 2028, None)]
GREY_PCT = dict(name="Arial", italic=True, theme=1, tint=0.35)


def edge(col, bottom=None):
    """분기 열(F·G) 양옆 세로선 + 행 아래선."""
    b = {}
    if col == "F":
        b["left"] = THIN
    if col == "G":
        b["right"] = THIN
    if bottom:
        b["bottom"] = bottom
    return b


def table_head(ws, r, title, band_theme):
    put(ws, f"A{r}", "X", kfont(), align="center")
    put(ws, f"B{r}", title, kfont("H", theme=0), fil=fill(band_theme))
    for col, *_ in YEARS:
        put(ws, f"{col}{r}", fnt=kfont(), fil=fill(band_theme))
    put(ws, f"B{r + 1}", "(단위: 백만 원)", kfont("B"), bottom=DBL, right=THIN)
    for col, head, actual in YEARS:
        put(ws, f"{col}{r + 1}", head, kfont("B"), nf=None if actual else 'General"E"', align="right",
            **edge(col, DBL))


def line(ws, r, label, formula, kind, bottom=None):
    """kind: total / yoy / item / child / pct.  formula(열, RAW 기간) -> 수식 (추정 열은 비워 둠)."""
    lab = {
        "total": dict(fnt=kfont("B"), fil=fill(3)),
        "yoy": dict(fnt=font(bold=True, italic=True, theme=4), nf=PCT, align="right"),
        "item": dict(fnt=kfont()),
        "child": dict(fnt=kfont(), align="left", indent=1),
        "pct": dict(fnt=font(**GREY_PCT), nf=PCT, align="right"),
    }[kind]
    side = {"right": THIN, **({"bottom": bottom} if bottom else {})}
    put(ws, f"B{r}", label, **lab, **side)
    for col, _, actual in YEARS:
        style = {
            "total": dict(fnt=font(bold=True, theme=1), nf=MN, fil=fill(3 if actual else 2)),
            "yoy": dict(fnt=font(bold=True, italic=True, theme=4), nf=PCT),
            "item": dict(fnt=font(theme=1), nf=MN),
            "child": dict(fnt=font(theme=1), nf=MN),
            "pct": dict(fnt=font(**GREY_PCT), nf=PCT),
        }[kind]
        put(ws, f"{col}{r}", formula(col, actual) if actual else None, **style, **edge(col, bottom))


def raw_ref(rows, key):
    return lambda col, period: f"={RAW}!{COL[period]}{rows[key]}"


def build_rev(wb, rows):
    ws = fresh_sheet(wb, REV)
    ws.column_dimensions["A"].width = 2.6
    ws.column_dimensions["B"].width = 18
    for col, *_ in YEARS:
        ws.column_dimensions[col].width = 10.6
    table_head(ws, 2, "최종 매출 추정 Table", 4)
    line(ws, 4, "매출액", raw_ref(rows, "tot"), "total")
    prev = {"D": "C", "E": "D"}  # 2024·2025는 같은 표의 전년 칸, 2023·분기는 RAW의 전년(동기) 칸

    def yoy(c, p):
        return f"={c}4/{prev[c]}4-1" if c in prev else f"={c}4/{RAW}!{COL[PRIOR[p]]}{rows['tot']}-1"

    line(ws, 5, "YoY (%)", yoy, "yoy", bottom=THIN)
    r = 6
    segs = [("comp", "컴포넌트"), ("pkg", "패키지솔루션"), ("opt", "광학솔루션")]
    for k, name in segs:
        last = k == "opt"
        line(ws, r, name, raw_ref(rows, k), "item")
        line(ws, r + 1, "% of sales", lambda c, p, r=r: f"={c}{r}/{c}$4", "pct")
        line(ws, r + 2, "수출", raw_ref(rows, f"{k}_exp"), "child")
        line(ws, r + 3, "내수", raw_ref(rows, f"{k}_dom"), "child", bottom=MED if last else None)
        r += 4


def build_est(wb, rows):
    ws = fresh_sheet(wb, EST)
    ws.column_dimensions["A"].width = 2.6
    ws.column_dimensions["B"].width = 28.3
    for col, *_ in YEARS:
        ws.column_dimensions[col].width = 10.6
    r = 2
    for i, (k, name) in enumerate([("comp", "컴포넌트"), ("pkg", "패키지솔루션"), ("opt", "광학솔루션")]):
        table_head(ws, r, f"{name} 매출 추정", 9)
        top = r + 2
        share = lambda c, p, row: f"={c}{row}/{c}${top}"
        line(ws, top, f"{name} 매출액", raw_ref(rows, k), "total")
        line(ws, top + 1, "% of sales", lambda c, p, top=top: f"={c}{top}/{REV}!{c}$4", "pct", bottom=THIN)
        items = [("수출", f"{k}_exp"), ("내수", f"{k}_dom"), ("제품 및 상품", f"{k}_goods"), ("기타 (용역 등)", f"{k}_other")]
        for j, (label, key) in enumerate(items):
            row = top + 2 + 2 * j
            line(ws, row, label, raw_ref(rows, key), "item")
            bottom = MED if j == 3 else (THIN if j == 1 else None)
            line(ws, row + 1, f"% of {name}", lambda c, p, row=row, top=top: f"={c}{row}/{c}${top}", "pct",
                 bottom=bottom)
        notes = ["수출 + 내수 = 제품 및 상품 + 기타 = 부문 매출 (같은 매출을 DART 공시 기준 두 가지로 나눈 것)."]
        for j, text in enumerate(notes):
            put(ws, f"A{top + 11 + j}", "-", kfont(), align="center")
            put(ws, f"B{top + 11 + j}", text, kfont())
        r += 17


def main(src, dst):
    wb = openpyxl.load_workbook(src)
    rows = build_raw(wb)
    build_rev(wb, rows)
    build_est(wb, rows)
    wb.calculation.fullCalcOnLoad = True
    wb.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
