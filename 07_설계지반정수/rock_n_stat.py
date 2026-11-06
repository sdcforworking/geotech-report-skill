"""암종별 N값 모음 엑셀: 오른쪽 지층별 표(P:R)에 30cm 환산 N값 최소·최대·범위·평균 열을 수식으로 추가.
사용: python rock_n_stat.py 입력.xlsx 출력.xlsx"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC, OUT = sys.argv[1:3]
wb = openpyxl.load_workbook(SRC)
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
for ws in wb:
    n = max(r for r in range(5, ws.max_row + 1) if ws.cell(r, 1).value is not None)
    E, Fc, M = f'$E$5:$E${n}', f'$F$5:$F${n}', f'$M$5:$M${n}'
    ws['R4'] = '시험\n횟수'
    for j, h in zip(range(19, 23), ['최소\n(환산 N)', '최대\n(환산 N)', '범위\n(환산 N)', '평균\n(환산 N)']):
        c = ws.cell(4, j, h); c.font = FB; c.fill = HF; c.border = BD; c.alignment = C
    k = 5
    while ws.cell(k, 16).value not in (None, '합계'):
        ws.cell(k, 19, f'=_xlfn.MINIFS({M},{E},P{k},{Fc},Q{k})')
        ws.cell(k, 20, f'=_xlfn.MAXIFS({M},{E},P{k},{Fc},Q{k})')
        ws.cell(k, 22, f'=AVERAGEIFS({M},{E},P{k},{Fc},Q{k})')
        k += 1
    # 합계 행: 암종 전체
    ws.cell(k, 19, f'=MIN({M})'); ws.cell(k, 20, f'=MAX({M})'); ws.cell(k, 22, f'=AVERAGE({M})')
    for r in range(5, k + 1):
        ws.cell(r, 21, f'=TEXT(S{r},"0.0")&"∼"&TEXT(T{r},"0.0")')
        for j in range(19, 23):
            c = ws.cell(r, j); c.font = FB if r == k else F; c.border = BD; c.alignment = C
            if j != 21:
                c.number_format = '0.0'
    for col, w in (('S', 9), ('T', 9), ('U', 13), ('V', 9)):
        ws.column_dimensions[col].width = w
    ws.cell(k + 1, 16, '※ 환산 N = 타격수 × 30 / 관입량 (M열). 50회 관입불능 값도 그대로 포함한 단순 최소·최대·평균').font = F
wb.save(OUT)
print('저장', OUT)
