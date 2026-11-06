"""4.2.2 정리 엑셀 '암종별 N값 자료'(720회)를 암 구분(화강암/편마암)별 시트 두 개로 나눠 새 엑셀로 만듦.
지층 순서(붕적층 모래→자갈→풍화토 N≤30→N＞30→풍화암) → 공번(보고서 순서) → 심도 순. 30cm 환산 N값은 엑셀 수식.
사용: python rock_n.py 출력.xlsx"""
import sys, glob, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC = glob.glob(r'C:\Users\심대천\Desktop\동해안-신가평 보고서\1. 최종 정리\05_보고서\4.2.2*.xlsx')[0]
OUT = sys.argv[1]
wb0 = openpyxl.load_workbook(SRC, data_only=True)

# 공번 순서·암 구분 (암종 구분 시트)
G, ORD = {}, {}
for r in wb0['암종 구분'].iter_rows(min_row=2, values_only=True):
    if r[1] and str(r[1]).startswith('No.'):
        ORD[r[1]] = len(ORD); G[r[1]] = r[6]

rows = []
for r in wb0['암종별 N값 자료'].iter_rows(min_row=2, values_only=True):
    if r[0] not in ('화강암', '편마암'):
        continue
    gu, am, tw, hole, ly, sub, comp, uscs, d, nv, n30, rd = r[:12]
    m = re.fullmatch(r'(\d+)/(\d+(?:\.\d+)?)', str(nv))
    blow, pen = int(m.group(1)), float(m.group(2))
    rows.append(dict(gu=gu, am=am, tw=tw, hole=hole, ly=ly, sub=sub or '-', comp=comp, uscs=uscs or '-', d=d, nv=nv,
                     blow=blow, pen=pen, n30=n30, rd=rd))
assert len(rows) == 720, len(rows)

LO = {('붕적층', '모래'): 0, ('붕적층', '자갈'): 1, ('풍화토', 'N≤30'): 2, ('풍화토', 'N＞30'): 3, ('풍화암', '-'): 4}
rows.sort(key=lambda x: (LO[(x['ly'], x['sub'])], ORD[x['hole']], x['d']))

HDR = ['번호', '철탑\n번호', '시추\n공번', '암종\n(연암 암종)', '지층', '구분', '구성성분', '통일분류\n(xsm)', '시험심도\n(GL.-m)',
       'N값\n(회/cm)', '타격수\n(회)', '관입량\n(cm)', '30cm 환산\nN값', '상대밀도']
W = [6, 9, 12, 14, 8, 8, 22, 9, 9, 10, 8, 8, 10, 10]
thin = Side(style='thin', color='808080')
BD = Border(left=thin, right=thin, top=thin, bottom=thin)
HF = PatternFill('solid', fgColor='D9E1F2')
F = Font(name='맑은 고딕', size=10)
FB = Font(name='맑은 고딕', size=10, bold=True)
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
L = Alignment(horizontal='left', vertical='center')

wb = openpyxl.Workbook(); wb.remove(wb.active)
for gu in ('화강암', '편마암'):
    R = [x for x in rows if x['gu'] == gu]
    ws = wb.create_sheet(f'{gu} N값')
    holes = {x['hole'] for x in R}
    ws['A1'] = f'{gu} 표준관입시험 N값 ({len(R)}회 · {len(holes)}공)'; ws['A1'].font = Font(name='맑은 고딕', size=13, bold=True)
    ws['A2'] = "※ 암 구분은 4.2.2 정리 엑셀 '암종 구분' 시트 기준 (편마암에 화강암질 편마암 포함). 30cm 환산 N값 = 타격수 × 30 / 관입량 (수식)"
    ws['A2'].font = F
    for j, (h, w) in enumerate(zip(HDR, W), 1):
        c = ws.cell(4, j, h); c.font = FB; c.fill = HF; c.border = BD; c.alignment = C
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.row_dimensions[4].height = 42
    for i, x in enumerate(R):
        r = 5 + i
        vals = [i + 1, x['tw'], x['hole'], x['am'], x['ly'], x['sub'], x['comp'], x['uscs'], x['d'], x['nv'], x['blow'], x['pen'],
                f'=K{r}*30/L{r}', x['rd']]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v); c.font = F; c.border = BD; c.alignment = L if j == 7 else C
        ws.cell(r, 9).number_format = '0.0'
        ws.cell(r, 12).number_format = '0'
        ws.cell(r, 13).number_format = '0.0'
    ws.auto_filter.ref = f'A4:N{4 + len(R)}'
    ws.freeze_panes = 'A5'
    # 지층별 개수 (오른쪽 참고표, 수식)
    ws['P4'] = '지층'; ws['Q4'] = '구분'; ws['R4'] = '시험 횟수'
    for c in ('P4', 'Q4', 'R4'):
        ws[c].font = FB; ws[c].fill = HF; ws[c].border = BD; ws[c].alignment = C
    k = 5
    for (ly, sub) in LO:
        if not any(x['ly'] == ly and x['sub'] == sub for x in R):
            continue
        ws.cell(k, 16, ly); ws.cell(k, 17, sub)
        ws.cell(k, 18, f'=COUNTIFS($E$5:$E${4 + len(R)},P{k},$F$5:$F${4 + len(R)},Q{k})')
        k += 1
    ws.cell(k, 16, '합계'); ws.cell(k, 18, f'=SUM(R5:R{k - 1})')
    for rr in range(5, k + 1):
        for cc in (16, 17, 18):
            c = ws.cell(rr, cc); c.font = FB if rr == k else F; c.border = BD; c.alignment = C
    for col, w in (('O', 3), ('P', 8), ('Q', 8), ('R', 10)):
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = 'portrait'; ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = '4:4'; ws.print_area = f'A1:N{4 + len(R)}'
wb.save(OUT)
print(OUT, {g: sum(1 for x in rows if x['gu'] == g) for g in ('화강암', '편마암')})
# 암 구분이 '암종 구분' 시트와 같은지
bad = [(x['hole'], x['gu'], G.get(x['hole'])) for x in rows if G.get(x['hole']) != x['gu']]
print('암 구분 불일치', bad[:5], len(bad))
