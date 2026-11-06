"""암종별 N값 모음 엑셀:
1) '대표 N값' 시트에서 사용자가 노란색(FFFF00)으로 칠한 방법 칸(L~O)을 읽어 Q열(적용 방법)에 그 방법 이름을 넣음 → S열 대표 N 계산.
2) 'N값 내부마찰각' 시트 추가 — 대표 N(S열)을 끌어와 보고서 5.3.1과 같은 4개 경험식(도로교·Dunham·Peck·Ohsaki)으로 φ 계산.
사용: python rock_n_phi.py 입력.xlsx 출력.xlsx"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
rep = wb['대표 N값']
COLM = {12: '단순 평균', 13: '50 상한 평균', 14: '관입불능 제외 평균', 15: '중앙값'}
ROWS = []
for r in range(6, 14):
    hit = [c for c in COLM if rep.cell(r, c).fill.fill_type and rep.cell(r, c).fill.fgColor.rgb == 'FFFFFF00']
    assert len(hit) <= 1, (r, hit)
    if hit:
        assert rep.cell(r, 17).value in (None, COLM[hit[0]]), (r, rep.cell(r, 17).value)
        rep.cell(r, 17, COLM[hit[0]])
        ROWS.append(r)
print('노란 표시 행', [(rep.cell(r, 1).value, rep.cell(r, 2).value, rep.cell(r, 3).value, rep.cell(r, 17).value) for r in ROWS])

thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
YF = PatternFill('solid', fgColor='FFF2CC')
GF = PatternFill('solid', fgColor='E2EFDA')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
L = Alignment(horizontal='left', vertical='center', wrap_text=True)

ws = wb.create_sheet('N값 내부마찰각', 1)
ws['A1'] = 'N값을 이용한 내부마찰각(φ) 산정 — 대표 N값은 \'대표 N값\' 시트에서 연결, 계산은 모두 함수'
ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
ws['A2'] = "※ 경험식은 보고서 5.3.1 '내부마찰각 산정식'과 같은 4개 (도로교설계기준 · Dunham · Peck · Ohsaki). 각 식 값은 소수 첫째 자리로 반올림 후 범위·평균"
ws['A3'] = "※ 노란 칸(Dunham 상수)만 직접 고름 — 아래 표 참고. 대표 N값을 바꾸려면 '대표 N값' 시트의 적용 방법(Q열)을 바꾸면 여기도 따라 바뀜"
for c in ('A2', 'A3'):
    ws[c].font = F
HDR = ['암 구분', '지층', '구분', '대표 N값\n(회/cm)', '적용 N\n(환산)', 'Dunham\n상수 C',
       '도로교\n√(15N)+15\n(≤45)', 'Dunham\n√(12N)+C', 'Peck\n0.3N+27', 'Ohsaki\n√(20N)+15',
       '최소', '최대', '범위', '평균']
W = [12, 8, 8, 10, 8, 9, 11, 10, 9, 10, 8, 8, 12, 8]
for j, (h, w) in enumerate(zip(HDR, W), 1):
    c = ws.cell(5, j, h); c.font = FB; c.border = BD; c.alignment = C
    c.fill = YF if j == 6 else GF if j >= 11 else HF
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.row_dimensions[5].height = 48
for j, t in ((1, '구분'), (4, '대표 N값 (연결)'), (7, '경험식 내부마찰각 φ (°)'), (11, '정리 (°)')):
    ws.cell(4, j, t).font = FB

for i, rr in enumerate(ROWS):
    r = 6 + i
    src = f"'대표 N값'!"
    ws.cell(r, 1, f'={src}A{rr}'); ws.cell(r, 2, f'={src}B{rr}'); ws.cell(r, 3, f'={src}C{rr}')
    ws.cell(r, 4, f'={src}T{rr}'); ws.cell(r, 5, f'={src}S{rr}')
    ws.cell(r, 6, 20)
    ok = f'ISNUMBER(E{r})'
    ws.cell(r, 7, f'=IF({ok},ROUND(MIN(SQRT(15*E{r})+15,45),1),"")')
    ws.cell(r, 8, f'=IF({ok},ROUND(SQRT(12*E{r})+F{r},1),"")')
    ws.cell(r, 9, f'=IF({ok},ROUND(0.3*E{r}+27,1),"")')
    ws.cell(r, 10, f'=IF({ok},ROUND(SQRT(20*E{r})+15,1),"")')
    ws.cell(r, 11, f'=IF({ok},MIN(G{r}:J{r}),"")')
    ws.cell(r, 12, f'=IF({ok},MAX(G{r}:J{r}),"")')
    ws.cell(r, 13, f'=IF({ok},TEXT(K{r},"0.0")&"∼"&TEXT(L{r},"0.0"),"")')
    ws.cell(r, 14, f'=IF({ok},AVERAGE(G{r}:J{r}),"")')
    for j in range(1, 15):
        c = ws.cell(r, j); c.font = F; c.border = BD; c.alignment = C
        if j == 6:
            c.fill = YF
        if j >= 7 and j != 13:
            c.number_format = '0.0'
last = 6 + len(ROWS) - 1

# 경험식 · Dunham 상수 안내표
n = last + 2
ws.cell(n, 1, '■ 내부마찰각 경험식').font = FB; n += 1
for j, h in enumerate(['식 이름', '식', '비고'], 1):
    c = ws.cell(n, [1, 4, 8][j - 1], h); c.font = FB; c.fill = HF
n += 1
for a, b, cc in [('도로교설계기준', 'φ = √(15N) + 15  (단, φ ≤ 45°)', 'N > 5 인 사질토에 적용'),
                 ('Dunham (1954)', 'φ = √(12N) + C', 'C = 15 · 20 · 25 (입자 모양 · 입도에 따라, 아래 표)'),
                 ('Peck (1974)', 'φ = 0.3N + 27', '사질토'),
                 ('Ohsaki (1959)', 'φ = √(20N) + 15', '사질토')]:
    ws.cell(n, 1, a).font = F; ws.cell(n, 4, b).font = F; ws.cell(n, 8, cc).font = F; n += 1
n += 1
ws.cell(n, 1, '■ Dunham 상수 C').font = FB; n += 1
for j, h in enumerate(['C', '입자 모양 · 입도'], 1):
    c = ws.cell(n, [1, 4][j - 1], h); c.font = FB; c.fill = HF
n += 1
for a, b in [(15, '입자가 둥글고 입도가 균등 (입도 나쁨)'),
             (20, '입자가 둥글고 입도가 양호, 또는 입자가 모나고 입도가 균등'),
             (25, '입자가 모나고 입도가 양호')]:
    ws.cell(n, 1, a).font = F; ws.cell(n, 4, b).font = F; n += 1
ws.cell(n, 1, '※ 기존 보고서 5.3.1 표는 모든 지층에 C = 20을 씀 (그래서 기본값 20). 자갈·모난 입자면 25도 검토').font = F
ws.freeze_panes = 'D6'
ws.sheet_properties.tabColor = 'F4B084'
ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 1
wb.save(OUT)
print('저장', OUT, '| 행', len(ROWS))
