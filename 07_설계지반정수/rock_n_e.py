"""암종별 N값 모음 엑셀에 'N값 변형계수' 시트 추가 — 대표 N(대표 N값 시트 S열)과 계수('경험식 계수' 시트 α·Bowles 흙 종류)로
보고서 5.3.1 표와 같은 5개 식(Schmertmann · 구조물기초(Bowles) · Hisatake · 도로교 · 한국지반공학회) 계산. 평균은 보고서처럼 최대값 제외.
사용: python rock_n_e.py 입력.xlsx 출력.xlsx"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
phi = wb['N값 내부마찰각']
ROWS = []
for r in range(6, 30):
    f = phi.cell(r, 1).value
    if not (isinstance(f, str) and f.startswith("='대표 N값'!A")):
        break
    ROWS.append(int(f.split('!A')[1]))
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
GF = PatternFill('solid', fgColor='E2EFDA')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)

ws = wb.create_sheet('N값 변형계수', 2)
ws['A1'] = "N값을 이용한 변형계수(E) 산정 — 대표 N값 · 계수는 다른 시트에서 연결, 계산은 모두 함수"
ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
ws['A2'] = ("※ 식은 보고서 5.3.1 'N값을 이용한 변형계수 산정' 표와 같은 5개 (Schmertmann · 구조물기초(Bowles) · Hisatake · 도로교 · 한국지반공학회). "
            "각 식 값은 소수 첫째 자리로 반올림")
ws['A3'] = ("※ 평균은 보고서처럼 '제안자별 산정결과의 최대값을 제외한 평균'. Schmertmann α · Bowles 흙 종류는 '경험식 계수' 시트에서 바꿈")
for c in ('A2', 'A3'):
    ws[c].font = F
HDR = ['암 구분', '지층', '구분', '대표 N값\n(회/cm)', '적용 N\n(환산)', 'Schmertmann\nα', 'Bowles\n흙 종류',
       'Schmertmann\nα·N', '구조물기초\n(Bowles)', 'Hisatake\n0.5N+7', '도로교\n2.8N', '한국지반공학회\n1.2N',
       '최소', '최대', '범위', '평균\n(최대값 제외)', '평균\n(전체, 참고)']
W = [12, 8, 8, 10, 8, 12, 12, 12, 11, 10, 9, 13, 8, 8, 13, 11, 10]
for j, (h, w) in enumerate(zip(HDR, W), 1):
    c = ws.cell(5, j, h); c.font = FB; c.border = BD; c.alignment = C
    c.fill = GF if j in (6, 7) or j >= 13 else HF
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.row_dimensions[5].height = 42
for j, t in ((1, '구분'), (4, '대표 N값 · 계수 (연결)'), (8, '경험식 변형계수 E (MPa)'), (13, '정리 (MPa)')):
    ws.cell(4, j, t).font = FB

for i, rr in enumerate(ROWS):
    r = 6 + i
    s = "'대표 N값'!"
    ws.cell(r, 1, f'={s}A{rr}'); ws.cell(r, 2, f'={s}B{rr}'); ws.cell(r, 3, f'={s}C{rr}')
    ws.cell(r, 4, f'={s}T{rr}'); ws.cell(r, 5, f'={s}S{rr}')
    ws.cell(r, 6, f"='경험식 계수'!G{r}"); ws.cell(r, 7, f"='경험식 계수'!H{r}")
    ok = f'ISNUMBER(E{r})'
    ws.cell(r, 8, f'=IF({ok},ROUND(F{r}*E{r},1),"")')
    ws.cell(r, 9, f'=IF({ok},ROUND(IF(G{r}="실트질 모래",0.3*(E{r}+6),IF(G{r}="점토질 모래",0.32*(E{r}+15),'
                  f'IF(G{r}="모래",0.5*(E{r}+15),IF(G{r}="자갈질 모래",1.2*(E{r}+6),NA())))),1),"")')
    ws.cell(r, 10, f'=IF({ok},ROUND(0.5*E{r}+7,1),"")')
    ws.cell(r, 11, f'=IF({ok},ROUND(2.8*E{r},1),"")')
    ws.cell(r, 12, f'=IF({ok},ROUND(1.2*E{r},1),"")')
    ws.cell(r, 13, f'=IF({ok},MIN(H{r}:L{r}),"")')
    ws.cell(r, 14, f'=IF({ok},MAX(H{r}:L{r}),"")')
    ws.cell(r, 15, f'=IF({ok},TEXT(M{r},"0.0")&"∼"&TEXT(N{r},"0.0"),"")')
    ws.cell(r, 16, f'=IF({ok},(SUM(H{r}:L{r})-N{r})/(COUNT(H{r}:L{r})-1),"")')
    ws.cell(r, 17, f'=IF({ok},AVERAGE(H{r}:L{r}),"")')
    for j in range(1, 18):
        c = ws.cell(r, j); c.font = F; c.border = BD; c.alignment = C
        if j in (6, 7):
            c.fill = GF
        if j >= 8 and j != 15:
            c.number_format = '0.0'
last = 6 + len(ROWS) - 1
n = last + 2
for t in ['■ 변형계수 경험식',
          '  Schmertmann : E = α · N (MPa) — α는 흙 종류별 (실트 0.4 · 세립∼중립 모래 0.7 · 조립 모래 1.0 · 자갈질 모래·자갈 1.2∼1.5)',
          '  구조물기초설계기준 (Bowles) : 실트질 모래 0.30(N+6) · 점토질 모래 0.32(N+15) · 모래 0.50(N+15) · 자갈질 모래 1.20(N+6)',
          '  Hisatake : Es = 0.5N + 7 (MPa)',
          '  도로교 설계기준 : Ed = 2.8N (MPa)',
          '  한국지반공학회 : Ed = 1.2N (MPa)',
          "※ 기존 보고서 표는 Schmertmann α를 모래 1.0 · 풍화토와 자갈 1.2로 썼음 (N=13→13.0, N=18→21.6). 이 시트는 '경험식 계수' 시트 값(옛 파일 방식 0.7 / 1.2)을 따름",
          '※ 다음 단계: 문헌 범위 (Das · Bowles · 서울시 지반분류 등) · 기존 사례 · 공내재하시험(화강암 풍화토 N＞30 4회)과 비교해 적용값 결정']:
    ws.cell(n, 1, t).font = FB if t.startswith('■') else F; n += 1
ws.freeze_panes = 'D6'
ws.sheet_properties.tabColor = 'F4B084'
ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 9
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
wb.active = 2
wb.save(OUT)
print('저장', OUT, '| 행', len(ROWS), wb.sheetnames)
