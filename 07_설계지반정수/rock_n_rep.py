"""암종별 N값 모음 엑셀에 '대표 N값' 시트 추가 — 토사층(붕적층·풍화토) 지층별로 여러 방법의 평균을 함수로 계산하고,
노란 칸(적용 방법·직접 입력·판단 근거)을 고르면 대표 N값(회/cm)·상대밀도가 나옴. 관입불능 = 타격수 50회 · 관입량 30cm 미만.
사용: python rock_n_rep.py 입력.xlsx 출력.xlsx [시험용: 방법들을 채움]"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1:3]
TEST = len(sys.argv) > 3
wb = openpyxl.load_workbook(SRC)
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
YF = PatternFill('solid', fgColor='FFF2CC')
GF = PatternFill('solid', fgColor='E2EFDA')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
L = Alignment(horizontal='left', vertical='center', wrap_text=True)

# 데이터 시트 범위
RNG = {}
for g in ('화강암', '편마암'):
    ws = wb[f'{g} N값']
    n = max(r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value is not None)
    RNG[g] = {c: f"'{g} N값'!${c}$5:${c}${n}" for c in 'EFJKLM'}

METHODS = ['단순 평균', '50 상한 평균', '관입불능 제외 평균', '중앙값', '직접 입력']
ROWS = [('화강암', '붕적층', '모래'), ('화강암', '붕적층', '자갈'), ('화강암', '풍화토', 'N≤30'), ('화강암', '풍화토', 'N＞30'),
        ('편마암', '붕적층', '자갈'), ('편마암', '풍화토', 'N≤30'), ('편마암', '풍화토', 'N＞30'),
        ('화강암+편마암', '붕적층', '자갈')]

ws = wb.create_sheet('대표 N값', 0)
ws['A1'] = '암종별 지층별 대표 N값 (토사층) — 계산은 모두 함수, 노란 칸만 직접 고름'
ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
ws['A2'] = ("※ 환산 N = 타격수 × 30 / 관입량. 관입불능 = 타격수 50회에 관입량 30cm 미만인 시험 (50/27, 50/12 등). "
            "50 상한 평균 = 환산 N이 50을 넘으면 50으로 보고 평균.")
ws['A3'] = ("※ 노란 칸: 적용 방법을 목록에서 고르면 대표 N값·상대밀도가 나옴. '직접 입력'을 고르면 R열 값을 씀. "
            "풍화암·연암은 대표 N값을 쓰지 않음 (공내재하·문헌·사례, RMR).")
for c in ('A2', 'A3'):
    ws[c].font = F
HDR = ['암 구분', '지층', '구분', '시험\n횟수', '관입불능\n횟수', '최소\n(회/cm)', '최대\n(회/cm)',
       '단순\n평균', '50 상한\n평균', '관입불능\n제외 평균', '중앙값',
       '단순\n평균', '50 상한\n평균', '관입불능\n제외 평균', '중앙값', '표준편차\n(환산)',
       '적용 방법\n(선택)', '직접 입력\nN', '대표 N\n(환산)', '대표 N값\n(회/cm)', '상대밀도', '판단 근거 (직접 적음)']
W = [12, 8, 8, 7, 9, 9, 9, 8, 9, 10, 8, 8, 9, 10, 8, 9, 17, 9, 8, 10, 10, 50]
IN, OUTC = (17, 18, 22), (19, 20, 21)
CMF = PatternFill('solid', fgColor='FCE4D6')
for j, (h, w) in enumerate(zip(HDR, W), 1):
    c = ws.cell(5, j, h); c.font = FB; c.border = BD; c.alignment = C
    c.fill = YF if j in IN else GF if j in OUTC else CMF if 12 <= j <= 15 else HF
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.row_dimensions[5].height = 42
# 4행: 열 묶음 안내 (병합 없이)
for j, t in ((1, '구분'), (4, '자료 (함수)'), (8, '방법별 평균 — 환산 N (함수)'), (12, '방법별 평균 — 원래 형식 회/cm (함수)'),
             (17, '판단 (직접 고름)'), (19, '결과 (함수)')):
    c = ws.cell(4, j, t); c.font = FB
CM = lambda x: f'IF(ISNUMBER({x}),IF(ROUND({x},0)<=50,ROUND({x},0)&"/30","50/"&ROUND(1500/{x},0)),"-")'


def parts(g, ly, sub):
    """조건식 조각들: 해당 암종(들) 시트별 (조건, M범위, J범위, K, L)"""
    gs = ('화강암', '편마암') if '+' in g else (g,)
    return [(f'({RNG[x]["E"]}="{ly}")*({RNG[x]["F"]}="{sub}")', RNG[x]) for x in gs]


for i, (g, ly, sub) in enumerate(ROWS):
    r = 6 + i
    P = parts(g, ly, sub)
    S = lambda expr: '+'.join(f'SUMPRODUCT({c}*{expr(R)})' for c, R in P)
    refusal = lambda R: f'({R["K"]}=50)*({R["L"]}<30)'
    cnt = S(lambda R: '1')
    nref = S(refusal)
    ws.cell(r, 1, g); ws.cell(r, 2, ly); ws.cell(r, 3, sub)
    ws.cell(r, 4, f'={cnt}')
    ws.cell(r, 5, f'={nref}')
    # 최소·최대 (회/cm): 환산 N 최소·최대인 시험의 표기
    if len(P) == 1:
        c, R = P[0]
        ws.cell(r, 6, f'=INDEX({R["J"]},MATCH(1,INDEX({c}*({R["M"]}=_xlfn.MINIFS({R["M"]},{R["E"]},B{r},{R["F"]},C{r})),0),0))')
        ws.cell(r, 7, f'=INDEX({R["J"]},MATCH(1,INDEX({c}*({R["M"]}=_xlfn.MAXIFS({R["M"]},{R["E"]},B{r},{R["F"]},C{r})),0),0))')
        ws.cell(r, 11, f'=_xlfn.AGGREGATE(17,6,{R["M"]}/({c}),2)')
    else:
        for j in (6, 7, 11):
            ws.cell(r, j, '-')
    ws.cell(r, 8, f'=({S(lambda R: R["M"])})/D{r}')
    ws.cell(r, 9, f'=({S(lambda R: "((" + R["M"] + ">50)*50+(" + R["M"] + "<=50)*" + R["M"] + ")")})/D{r}')
    ws.cell(r, 10, f'=IF(D{r}=E{r},"-",({S(lambda R: "(1-" + refusal(R) + ")*" + R["M"])})/(D{r}-E{r}))')
    for j, src in zip(range(12, 16), 'HIJK'):
        ws.cell(r, j, f'={CM(src + str(r))}')
    ws.cell(r, 16, f'=IF(D{r}<2,"-",SQRT(({S(lambda R: "(" + R["M"] + f"-H{r})^2")})/(D{r}-1)))')
    ws.cell(r, 19, f'=IF(Q{r}="","",IF(Q{r}="직접 입력",IF(R{r}="","",ROUND(R{r},0)),'
                   f'ROUND(CHOOSE(MATCH(Q{r},{{"단순 평균","50 상한 평균","관입불능 제외 평균","중앙값"}},0),H{r},I{r},J{r},K{r}),0)))')
    ws.cell(r, 20, f'=IF(S{r}="","",IF(S{r}>50,"50/"&ROUND(1500/S{r},0),S{r}&"/30"))')
    ws.cell(r, 21, f'=IF(S{r}="","",IF(S{r}<=4,"매우느슨",IF(S{r}<=10,"느슨",IF(S{r}<=30,"보통조밀",IF(S{r}<=50,"조밀","매우조밀")))))')
    if TEST:
        ws.cell(r, 17, METHODS[i % 5]); ws.cell(r, 18, 40 if METHODS[i % 5] == '직접 입력' else None)
    for j in range(1, 23):
        c = ws.cell(r, j); c.font = F; c.border = BD; c.alignment = L if j == 22 else C
        if j in IN:
            c.fill = YF
        if 8 <= j <= 11 or j == 16:
            c.number_format = '0.0'
last = 6 + len(ROWS) - 1
dv = DataValidation(type='list', formula1='"' + ','.join(METHODS) + '"', allow_blank=True)
dv.add(f'Q6:Q{last}'); ws.add_data_validation(dv)
n = last + 2
for t in ["※ '화강암+편마암 붕적층 자갈' 행은 참고용 — 편마암 자갈이 6회뿐이라 두 암종을 합친 값과 비교해 볼 때 씀 (중앙값·최소·최대는 계산 안 함).",
          "※ 원래 형식(회/cm): 환산 N을 반올림해 50 이하면 \"N/30\", 50을 넘으면 타격수 50회로 보고 \"50/관입량\"(관입량 = 1500 ÷ 환산 N)",
          "※ 상대밀도 (환산 N): 4 이하 매우느슨 · 4∼10 느슨 · 10∼30 보통조밀 · 30∼50 조밀 · 50 초과 매우조밀",
          "※ 대표 N값은 다음 단계의 내부마찰각(φ)·변형계수(E) 경험식에 들어감 → 크게 잡으면 위험측"]:
    ws.cell(n, 1, t).font = F; n += 1
ws.freeze_panes = 'D6'
ws.sheet_properties.tabColor = 'F4B084'
ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 0
wb.save(OUT)
print('저장', OUT, '| 행', len(ROWS))
