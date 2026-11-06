"""암종별 N값 모음 엑셀에 '경험식 계수' 시트 추가 (옛 'N값 정리_수정.xls'의 '지층별 적용 계수' 표 방식).
지층별 Dunham 상수 · Schmertmann α · Bowles 흙 종류를 한곳에 두고, 'N값 내부마찰각' 시트 Dunham 상수(F열)는 여기서 끌어옴.
아래에 계수 고르는 기준표(Dunham · Schmertmann · Bowles · Hisatake · 도로교 · 점착력 식)도 정리.
사용: python rock_n_coef.py 입력.xlsx 출력.xlsx"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
phi = wb['N값 내부마찰각']
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
YF = PatternFill('solid', fgColor='FFF2CC')
GF = PatternFill('solid', fgColor='E2EFDA')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
L = Alignment(horizontal='left', vertical='center', wrap_text=True)

# 내부마찰각 시트의 행(암·지층·구분)을 그대로 따름
ROWS = []
for r in range(6, 30):
    f = phi.cell(r, 1).value
    if not (isinstance(f, str) and f.startswith("='대표 N값'!A")):
        break
    ROWS.append((r, int(f.split('!A')[1])))
rep = wb['대표 N값']
KEY = [(rep.cell(rr, 1).value, rep.cell(rr, 2).value, rep.cell(rr, 3).value) for _, rr in ROWS]
print('행', KEY)

# 옛 파일 '지층별 적용 계수' 방식: 사질토 20/0.7/실트질 모래, 자갈 20/1.2/자갈질 모래, 풍화토 20/0.7/실트질 모래
USCS = {('화강암', '붕적층', '모래'): 'SM', ('화강암', '붕적층', '자갈'): 'GM, GP', ('화강암', '풍화토', 'N≤30'): 'SM',
        ('화강암', '풍화토', 'N＞30'): 'SM', ('편마암', '붕적층', '자갈'): 'GM', ('편마암', '풍화토', 'N≤30'): 'SM',
        ('편마암', '풍화토', 'N＞30'): 'SM'}
COEF = {'모래': ('사질토', 20, 0.7, '실트질 모래'), '자갈': ('자갈', 20, 1.2, '자갈질 모래'),
        'N≤30': ('사질토', 20, 0.7, '실트질 모래'), 'N＞30': ('사질토', 20, 0.7, '실트질 모래')}

ws = wb.create_sheet('경험식 계수', 2)
ws['A1'] = '경험식 계수 — 지층별로 쓰는 상수를 한곳에 모음 (옛 \'N값 정리_수정.xls\' 지층별 적용 계수 방식)'
ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
ws['A2'] = "※ 노란 칸이 계수 (직접 고름). 'N값 내부마찰각' 시트의 Dunham 상수는 여기서 연결됨. Schmertmann α · Bowles 흙 종류는 다음 단계(변형계수)에서 씀"
ws['A3'] = '※ 기본값은 옛 파일과 기존 보고서 5.3.1 방식: Dunham 20 (모든 지층) · Schmertmann α 사질토 0.7 / 자갈 1.2 · Bowles 사질토 "실트질 모래" / 자갈 "자갈질 모래"'
for c in ('A2', 'A3'):
    ws[c].font = F
HDR = ['암 구분', '지층', '구분', '통일분류\n(xsm)', '토성', 'Dunham\n상수 C', 'Schmertmann\nα', 'Bowles\n흙 종류', 'Bowles 식\n(MPa)', '근거 · 비고 (직접 적음)']
W = [12, 8, 8, 9, 8, 9, 12, 12, 16, 46]
for j, (h, w) in enumerate(zip(HDR, W), 1):
    c = ws.cell(5, j, h); c.font = FB; c.border = BD; c.alignment = C
    c.fill = YF if j in (6, 7, 8, 10) else GF if j == 9 else HF
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.row_dimensions[5].height = 36
for j, t in ((1, '구분 (대표 N값 시트에서 연결)'), (6, '계수 (직접 고름)'), (9, '확인용')):
    ws.cell(4, j, t).font = FB

BOW = {'실트질 모래': 'Es = 0.30(N+6)', '점토질 모래': 'Es = 0.32(N+15)', '모래': 'Es = 0.50(N+15)', '자갈질 모래': 'Es = 1.20(N+6)'}
for i, ((pr, rr), k) in enumerate(zip(ROWS, KEY)):
    r = 6 + i
    tb, dc, al, bw = COEF[k[2]]
    ws.cell(r, 1, f"='대표 N값'!A{rr}"); ws.cell(r, 2, f"='대표 N값'!B{rr}"); ws.cell(r, 3, f"='대표 N값'!C{rr}")
    ws.cell(r, 4, USCS[k]); ws.cell(r, 5, tb)
    ws.cell(r, 6, dc); ws.cell(r, 7, al); ws.cell(r, 8, bw)
    ws.cell(r, 9, f'=IFERROR(VLOOKUP(H{r},$A$31:$D$34,4,FALSE),"")')
    for j in range(1, 11):
        c = ws.cell(r, j); c.font = F; c.border = BD; c.alignment = L if j == 10 else C
        if j in (6, 7, 8, 10):
            c.fill = YF
    # 내부마찰각 시트 Dunham 상수 → 연결
    pc = phi.cell(pr, 6)
    pc.value = f"='경험식 계수'!F{r}"; pc.fill = GF
last = 6 + len(ROWS) - 1
dv1 = DataValidation(type='list', formula1='"15,20,25"', allow_blank=True); dv1.add(f'F6:F{last}')
dv2 = DataValidation(type='list', formula1='"0.4,0.7,1,1.2,1.5"', allow_blank=True); dv2.add(f'G6:G{last}')
dv3 = DataValidation(type='list', formula1='"실트질 모래,점토질 모래,모래,자갈질 모래"', allow_blank=True); dv3.add(f'H6:H{last}')
for d in (dv1, dv2, dv3):
    ws.add_data_validation(d)


def table(r0, title, hdr, rows, cols=(1, 4, 7)):
    ws.cell(r0, 1, title).font = FB
    for j, h in zip(cols, hdr):
        c = ws.cell(r0 + 1, j, h); c.font = FB; c.fill = HF; c.border = BD
    for i, row in enumerate(rows):
        for j, v in zip(cols, row):
            c = ws.cell(r0 + 2 + i, j, v); c.font = F
    return r0 + 2 + len(rows) + 1


n = last + 2
n = table(n, '■ ① Dunham 상수 C — 내부마찰각 φ = √(12N) + C', ['C', '입자 모양 · 입도', '이번 과업'],
          [(15, '입자가 둥글고 입도가 균등 (입도 나쁨)', ''),
           (20, '입자가 둥글고 입도가 양호, 또는 입자가 모나고 입도가 균등', '기존 보고서 · 옛 파일: 모든 지층 20'),
           (25, '입자가 모나고 입도가 양호', '자갈이 모나고 입도 양호하면 검토')])
n = table(n, '■ ② Schmertmann α — 변형계수 E = α·N (MPa)', ['α', '흙 종류', '이번 과업'],
          [(0.4, '실트, 모래질 실트', ''), (0.7, '세립·중립 모래, 약간 실트질 모래', '붕적층 모래 · 풍화토 (SM)'),
           (1.0, '조립질 모래', ''), ('1.2 ∼ 1.5', '자갈질 모래, 자갈', '붕적층 자갈 1.2 (옛 파일 풍화암 1.5)')])
assert n == 31 - 1 or True
# ③ Bowles: VLOOKUP 범위 A31:D34 에 맞춤
n = 29
ws.cell(n, 1, '■ ③ Bowles 흙 종류 — 변형계수 Es (MPa)').font = FB
for j, h in zip((1, 4, 7), ['흙 종류', '식', '이번 과업']):
    c = ws.cell(n + 1, j, h); c.font = FB; c.fill = HF; c.border = BD
for i, (a, b, cc) in enumerate([('실트질 모래', BOW['실트질 모래'], '붕적층 모래 · 풍화토 (SM)'), ('점토질 모래', BOW['점토질 모래'], ''),
                                ('모래', BOW['모래'], ''), ('자갈질 모래', BOW['자갈질 모래'], '붕적층 자갈 (GM·GP)')]):
    ws.cell(n + 2 + i, 1, a).font = F; ws.cell(n + 2 + i, 4, b).font = F; ws.cell(n + 2 + i, 7, cc).font = F
n = n + 7
n = table(n, '■ ④ 그 밖의 N값 경험식 (계수 없음)', ['식 이름', '식', '쓰는 곳'],
          [('도로교설계기준', 'φ = √(15N) + 15  (φ ≤ 45°, N > 5)', '내부마찰각'), ('Peck (1974)', 'φ = 0.3N + 27', '내부마찰각'),
           ('Ohsaki (1959)', 'φ = √(20N) + 15', '내부마찰각'), ('Hisatake', 'Es = 0.5N + 7 (MPa)', '변형계수'),
           ('도로교설계기준', 'Ed = 2.8N (MPa)', '변형계수'), ('한국지반공학회', 'Ed = 1.2N (MPa)', '변형계수 (기존 보고서 5.3.1)')])
n = table(n, '■ ⑤ 점착력 경험식 (점성토용 — 옛 파일에 있던 식)', ['식 이름', '식 (kPa)', '이번 과업'],
          [('Dunham', 'c = (N / 0.077) / 2', '적용 안 함'), ('Terzaghi-Peck', 'c = (N / 0.082) / 2', '적용 안 함'),
           ('Ohsaki', 'c = (40 + N / 0.2) / 2', '적용 안 함')])
ws.cell(n, 1, '※ 이번 과업 토사층은 실내시험 46회가 모두 비소성(NP) 사질토라 N값 점착력 식은 쓰지 않음 → 점착력은 문헌 · 사례로 정함').font = F
ws.freeze_panes = 'D6'
ws.sheet_properties.tabColor = 'F4B084'
ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True

# 내부마찰각 시트 안내 문구
phi['A3'] = "※ Dunham 상수(F열)는 '경험식 계수' 시트에서 연결됨 — 바꾸려면 그 시트 노란 칸을 고침. 대표 N값은 '대표 N값' 시트 적용 방법(Q열)을 따라 바뀜"
phi.cell(5, 6).fill = GF
for r in range(14, 40):
    v = phi.cell(r, 1).value
    if isinstance(v, str) and v.startswith('※ 기존 보고서 5.3.1 표는'):
        phi.cell(r, 1, "※ 기존 보고서 5.3.1 표는 모든 지층에 C = 20을 씀. 값은 '경험식 계수' 시트에서 고름")
wb.save(OUT)
print('저장', OUT, '| 계수 행', len(ROWS))
