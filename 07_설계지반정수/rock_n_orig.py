"""암종별 N값 모음 엑셀: 지층별 표에 원래 형식(회/cm) 최소·최대·범위·평균 열(W:Z)을 수식으로 추가.
최소·최대 = 환산 N 최소·최대가 나온 시험의 N값 표기(J열). 평균 = 환산 평균이 50 이하면 반올림/30, 넘으면 50/(1500÷평균 반올림).
사용: python rock_n_orig.py 입력.xlsx 출력.xlsx"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='FCE4D6')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
for ws in wb:
    n = max(r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value is not None)
    E, Fc, M, J = (f'${c}$5:${c}${n}' for c in 'EFMJ')
    for j, h in zip(range(23, 27), ['최소\n(회/cm)', '최대\n(회/cm)', '범위\n(회/cm)', '평균\n(회/cm)']):
        c = ws.cell(4, j, h); c.font = FB; c.fill = HF; c.border = BD; c.alignment = C
    k = 5
    while ws.cell(k, 16).value != '합계':
        ws.cell(k, 23, f'=INDEX({J},MATCH(1,INDEX(({E}=P{k})*({Fc}=Q{k})*({M}=S{k}),0),0))')
        ws.cell(k, 24, f'=INDEX({J},MATCH(1,INDEX(({E}=P{k})*({Fc}=Q{k})*({M}=T{k}),0),0))')
        k += 1
    ws.cell(k, 23, f'=INDEX({J},MATCH(S{k},{M},0))')
    ws.cell(k, 24, f'=INDEX({J},MATCH(T{k},{M},0))')
    for r in range(5, k + 1):
        ws.cell(r, 25, f'=W{r}&"∼"&X{r}')
        ws.cell(r, 26, f'=IF(ROUND(V{r},0)<=50,ROUND(V{r},0)&"/30","50/"&ROUND(1500/V{r},0))')
        for j in range(23, 27):
            c = ws.cell(r, j); c.font = FB if r == k else F; c.border = BD; c.alignment = C
    for col, w in (('W', 9), ('X', 9), ('Y', 14), ('Z', 9)):
        ws.column_dimensions[col].width = w
    ws.cell(k + 2, 16, '※ 회/cm 최소·최대: 환산 N이 최소·최대인 시험의 N값 표기. 평균(회/cm): 환산 평균이 50 이하면 "평균/30", 50 넘으면 "50/관입량"(관입량 = 1500 ÷ 평균)').font = F
wb.save(OUT)
print('저장', OUT)
