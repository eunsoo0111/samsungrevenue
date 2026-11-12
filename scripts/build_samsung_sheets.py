"""삼성전기 매출 시트 3장(삼성전기_RAW / 삼성전기_REV / 삼성전기_매출추정)을 다시 만든다.

사용법:  python3 scripts/build_samsung_sheets.py 입력.xlsx 출력.xlsx

- RAW 숫자는 DART 정기보고서(사업의 내용 II-2·3·4, 연결주석 3·20-1)에서 옮긴 값이다.
- 금액은 원 단위로 저장하고 표시형식(#,##0,,)으로 백만 원을 보여준다.
- REV·매출추정의 실적 칸은 전부 삼성전기_RAW 셀 하나를 그대로 참조한다.
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color

RAW, REV, EST = "삼성전기_RAW", "삼성전기_REV", "삼성전기_매출추정"

# ── RAW 열 배치 (닌텐도 R_REV와 동일: 연간 4개 + 분기 2개 + 빈 열 + 전년 동기 2개)
PERIODS = ["2022", "2023", "2024", "2025", "1Q26", "2Q26", "1Q25", "2Q25"]
COL = dict(zip(PERIODS, "CDEFGHJK"))
PRIOR = {"2023": "2022", "2024": "2023", "2025": "2024", "1Q26": "1Q25", "2Q26": "2Q25"}
HEAD = {
    "2022": ("제50기", "FY24 사업보고서 (재작성)"),
    "2023": ("제51기", "FY25 사업보고서 (재작성)"),
    "2024": ("제52기", "FY25 사업보고서"),
    "2025": ("제53기", "FY25 사업보고서"),
    "1Q26": ("제54기 1분기 (3M)", "1Q26 분기보고서"),
    "2Q26": ("제54기 2분기 (3M)", "1H26 반기 − 1Q26"),
    "1Q25": ("제53기 1분기 (3M)", "1Q25 분기보고서"),
    "2Q25": ("제53기 2분기 (3M)", "1H25 반기 − 1Q25"),
}
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
# 2. 매출유형별 — 부문별: 연결주석 20-1 (분기 열은 3개월 공시값), 합계: II-4 매출 유형별 매출실적
TYPE = {
    "comp_goods": [N, 3905722, 4443098, 5194938, 1405602, 1643792, 1199835, 1272571],
    "comp_other": [N, -2700, 18952, 3563, 2896, 5629, 16431, 8145],
    "pkg_goods": [N, 1460016, 1680718, 1999104, 637832, 699572, 426565, 486809],
    "pkg_other": [N, 257362, 353939, 302649, 87186, 72059, 72837, 77764],
    "opt_goods": [N, 3252588, 3747133, 3785921, 1070374, 1028428, 1015179, 926725],
    "opt_other": [N, 19424, 50263, 28284, 5223, 7725, 7802, 12615],
    "tot_goods": [9255753, 8618326, 9870949, 10979963, 3113808, 3371792, 2641579, 2686105],
    "tot_other": [150857, 274086, 423154, 334496, 95305, 85413, 97070, 98524],
}
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
def build_raw(wb):
    ws = fresh_sheet(wb, RAW)
    for col, width in {"A": 1.6, "B": 70, "I": 2.5}.items():
        ws.column_dimensions[col].width = width
    for col in "CDEFGHJK":
        ws.column_dimensions[col].width = 14.8

    rows = {}  # key -> 행 번호 (REV·매출추정이 참조)
    state = {"r": 2}

    def nxt():
        state["r"] += 1
        return state["r"]

    def literal(v, scale):
        if v is None:
            return None
        if isinstance(v, tuple):
            return f"={v[0] * scale}-{v[1] * scale}"
        return v * scale

    def section(title):
        r = nxt() + 1
        state["r"] = r
        put(ws, f"A{r}", "X", font(bold=True))
        put(ws, f"B{r}", title, font(bold=True))

    def data(key, label, values, nf=MN, scale=10**6, bold=False, indent=0):
        r = nxt()
        rows[key] = r
        put(ws, f"B{r}", label, font(bold=bold), indent=indent)
        for p, v in zip(PERIODS, values):
            put(ws, f"{COL[p]}{r}", literal(v, scale), font(theme=1), nf=nf)
        return r

    def calc(key, label, formula, nf, bold=True, periods=PERIODS, last=False):
        """formula(열 문자, 기간) -> 수식 문자열"""
        r = nxt()
        rows[key] = r
        side = {"bottom": THIN} if last else {}
        put(ws, f"B{r}", label, font(bold=True), **side)
        for p in PERIODS:
            f = formula(COL[p], p) if p in periods else None
            put(ws, f"{COL[p]}{r}", f, font(bold=bold, theme=1), nf=nf, **side)
        return r

    def ratio(key, label, num, den, **kw):
        return calc(key, label, lambda c, p: f"={c}{rows[num]}/{c}{rows[den]}", PCT, **kw)

    def yoy(key, label, base, **kw):
        return calc(key, label, lambda c, p: f"={c}{rows[base]}/{COL[PRIOR[p]]}{rows[base]}-1", PCT,
                    periods=list(PRIOR), **kw)

    def check(key, label, parts, total, periods=PERIODS):
        return calc(key, label, lambda c, p: "=" + "+".join(f"{c}{rows[k]}" for k in parts) + f"-{c}{rows[total]}",
                    MN1, bold=False, periods=periods)

    # 머리글
    put(ws, "A2", "X", font(bold=True))
    put(ws, "B2", "매출 Breakdown (3개 사업부문 — 부문·수출/내수·매출유형·지역·주요 매출처·가격·생산 공시)",
        font(size=12, bold=True))
    put(ws, "J2", "(ref) 전년 동기 — 분기 YoY 계산용", font(size=9, italic=True, rgb=GREY))
    put(ws, "B3", "(단위: 백만 원)", font(bold=True), bottom=DBL)
    put(ws, "B4", "기수 (12월 결산)", font())
    put(ws, "B5", "Source (DART)", font(size=9, italic=True, rgb=GREY))
    for p in PERIODS:
        c = COL[p]
        put(ws, f"{c}3", int(p) if p.isdigit() else p, font(bold=True), align="right", bottom=DBL)
        put(ws, f"{c}4", HEAD[p][0], font(), align="right")
        put(ws, f"{c}5", HEAD[p][1], font(size=9, italic=True, rgb=GREY), align="right")
    state["r"] = 4  # 다음 section()이 6행에서 시작

    seg = [("comp", "컴포넌트", "수동소자: MLCC, Inductor, Chip Resistor 등"),
           ("pkg", "패키지솔루션", "반도체패키지기판"),
           ("opt", "광학솔루션", "카메라모듈")]

    section("1. 사업부문별 매출실적 (사업보고서 II-4 가. 매출실적)")
    for k, name, product in seg:
        data(f"{k}_exp", f"{name} — 수출", SALES[f"{k}_exp"], indent=1)
        data(f"{k}_dom", f"{name} — 내수", SALES[f"{k}_dom"], indent=1)
        data(k, f"{name} — 합계 ({product})", SALES[k], bold=True)
    data("tot_exp", "합계 — 수출", SALES["tot_exp"], indent=1)
    data("tot_dom", "합계 — 내수", SALES["tot_dom"], indent=1)
    data("tot", "매출액 합계 (as disclosed)", SALES["tot"], bold=True)
    check("chk_seg", "(check) Σ 부문 합계 − 매출액 합계", ["comp", "pkg", "opt"], "tot")
    check("chk_exp", "(check) 수출 + 내수 − 매출액 합계", ["tot_exp", "tot_dom"], "tot")
    for k, name, _ in seg:
        ratio(f"mix_{k}", f"Mix: {name} (% of total)", k, "tot")
    ratio("mix_exp", "Mix: 수출 (% of total)", "tot_exp", "tot")
    for k, name, _ in seg:
        yoy(f"yoy_{k}", f"YoY: {name} (%)", k)
    yoy("yoy_tot", "YoY: 매출액 합계 (%)", "tot", last=True)

    section("2. 매출유형별 매출 — 제품·상품 / 기타(용역 등) (부문별: 연결주석 20-1, 합계: 사업보고서 II-4)")
    for k, name, _ in seg:
        data(f"{k}_goods", f"{name} — 제품 및 상품", TYPE[f"{k}_goods"], indent=1)
        data(f"{k}_other", f"{name} — 기타 (용역 등)", TYPE[f"{k}_other"], indent=1)
    data("tot_goods", "합계 — 제품/상품", TYPE["tot_goods"], bold=True)
    data("tot_other", "합계 — 용역 및 기타매출", TYPE["tot_other"], bold=True)
    for k, name, _ in seg:
        check(f"chk_type_{k}", f"(check) {name} 제품·상품 + 기타 − 부문 합계 (±1 = 공시 반올림)",
              [f"{k}_goods", f"{k}_other"], k, periods=PERIODS[1:])
    check("chk_type", "(check) Σ 유형 − 매출액 합계", ["tot_goods", "tot_other"], "tot")
    for k, name, _ in seg:
        ratio(f"mix_other_{k}", f"Mix: {name} 기타 (% of 부문)", f"{k}_other", k,
              periods=PERIODS[1:], last=(k == "opt"))

    section("3. 지역별 매출 — 전사 기준 (연결주석 3. 영업부문정보 > 지역에 대한 공시, 부문별 지역 매출은 미공시)")
    regions = [("kr_dom", "국내 — 내수"), ("kr_exp", "국내 — 수출"), ("cn", "중국"), ("sea", "동남아"),
               ("am", "미주"), ("eu", "유럽"), ("jp", "일본")]
    for k, name in regions:
        data(k, name, REGION[k])
    check("chk_region", "(check) Σ 지역 − 매출액 합계", [k for k, _ in regions], "tot",
          periods=PERIODS[1:])
    for k, name in regions:
        ratio(f"mix_{k}", f"Mix: {name} (% of total)", k, "tot", periods=PERIODS[1:], last=(k == "jp"))

    section("4. 주요 매출처 — 삼성전자와 그 종속기업 (연결주석 3 > 주요 고객에 대한 공시, 억원 단위 공시)")
    data("sec", "삼성전자와 그 종속기업 매출", SEC)
    ratio("mix_sec", "Mix: 삼성전자와 그 종속기업 (% of total)", "sec", "tot", periods=PERIODS[1:], last=True)

    section("5. 평균판매가격 변동률 — 전년 대비 (사업보고서 II-2 나. 가격변동추이, 2Q 열은 반기 누적 기준)")
    data("asp_comp", "MLCC (컴포넌트)", ASP["comp"], nf=PCT, scale=1)
    data("asp_pkg", "반도체패키지기판 (패키지솔루션)", ASP["pkg"], nf=PCT, scale=1)
    data("asp_opt", "카메라모듈 (광학솔루션)", ASP["opt"], nf=PCT, scale=1)

    section("6. 생산능력·생산실적·가동률 (사업보고서 II-3 다·라, 글로벌 완제품 기준, 공시 단위 그대로)")
    for k, name, unit in [("comp", "컴포넌트", "억개"), ("pkg", "패키지솔루션", "천㎡"), ("opt", "광학솔루션", "백만개")]:
        data(f"{k}_cap", f"{name} — 생산능력 ({unit})", PROD[f"{k}_cap"], nf=QTY, scale=1, indent=1)
        data(f"{k}_out", f"{name} — 생산실적 ({unit})", PROD[f"{k}_out"], nf=QTY, scale=1, indent=1)
        ratio(f"{k}_util", f"{name} — 가동률 (calc. = 생산실적 ÷ 생산능력)", f"{k}_out", f"{k}_cap", last=(k == "opt"))

    r = nxt() + 1
    for i, text in enumerate([
        "금액은 원 단위로 저장하고 표시형식으로 백만 원을 보여줌. 2Q 열 = 반기 누적 − 1분기 (셀 수식에 두 숫자 그대로).",
        "2022·2023년은 중단사업 발생으로 재작성된 수치 (이후 보고서 기준). 2022년의 부문별 매출유형·지역·주요 매출처는 재작성 기준 미공시라 공란.",
        "광학솔루션은 2023년 보고서까지 '광학통신솔루션'으로 표기.",
    ]):
        put(ws, f"A{r + i}", "-", font())
        put(ws, f"B{r + i}", text, font())
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
    notes = ["실적(2023~2Q26)은 전부 삼성전기_RAW 1번 표를 그대로 링크. 2023·분기 YoY의 전년 값도 RAW (분기는 전년 동기 열).",
             "2026E~2028E는 추정 단계에서 입력."]
    for i, text in enumerate(notes):
        put(ws, f"A{r + 1 + i}", "-", kfont(), align="center")
        put(ws, f"B{r + 1 + i}", text, kfont())


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
        notes = ["수출 + 내수 = 제품 및 상품 + 기타 = 부문 매출 (같은 매출을 DART 공시 기준 두 가지로 나눈 것).",
                 "수출/내수: 삼성전기_RAW 1번 표, 제품 및 상품/기타: 삼성전기_RAW 2번 표."]
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
