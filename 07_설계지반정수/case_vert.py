"""'기존 설계사례 요약' 시트를 본문 5.2 표와 같은 배치로 다시 만듦 — 암종별로 행 = 정수(단위중량·점착력·내부마찰각), 열 = 표토/붕적층 모래·자갈, 풍화토.
값은 사용자가 준 본문 표 그대로. 아래에 같은 배치로 비교용 최소·최대 숫자 표. 병합 없음, 탭 하나만 선택.
사용: python case_vert.py 파일.xlsx"""
import os, sys, re
import win32com.client as w

P = os.path.abspath(sys.argv[1])
COLS = ['표토/붕적층\n모래', '표토/붕적층\n자갈', '풍화토']
PAR = ['단위중량(kN/m³)', '점착력(kPa)', '내부마찰각(°)']
DATA = {'화강암': [['18.0∼19.0', '19.0∼20.0', '18.0∼19.5'], ['0~5', '0', '10~25'], ['25~28', '32~38', '28~31']],
        '편마암': [['17.0∼18.5', '19.0', '18.0∼19.5'], ['0~5', '0', '10~25'], ['25~28.5', '30', '29~30']]}
mm = lambda t: (float(re.split(r'[~∼]', t)[0]), float(re.split(r'[~∼]', t)[-1]))
BLUE, GRAY = 0xF2E1D9, 0xEFEFEF

xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    old = wb.Worksheets('기존 설계사례 요약')
    prev = old.Index
    ws = wb.Worksheets.Add(None, old)
    old.Delete()
    ws.Name = '기존 설계사례 요약'
    ws.Cells.Font.Name = '맑은 고딕'; ws.Cells.Font.Size = 10
    ws.Range('A1').Value = '기존 설계자료 검토 요약 (본문 5.2 표와 같은 배치 · 값 그대로)'
    ws.Range('A1').Font.Size = 13; ws.Range('A1').Font.Bold = True

    def box(r0, c0, nr, nc):
        rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + nr - 1, c0 + nc - 1))
        for b in (7, 8, 9, 10, 11, 12):
            rng.Borders(b).LineStyle = 1; rng.Borders(b).Weight = 2; rng.Borders(b).Color = 0x808080

    def head(r, labels):
        for j, h in enumerate(['구  분'] + labels, 1):
            c = ws.Cells(r, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = BLUE
            c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
        ws.Rows(r).RowHeight = 30

    r = 3
    for g, rows in DATA.items():
        ws.Cells(r, 1).Value = f'■ {g}'; ws.Cells(r, 1).Font.Bold = True
        head(r + 1, COLS)
        for i, (p, vals) in enumerate(zip(PAR, rows)):
            ws.Cells(r + 2 + i, 1).Value = p
            for j, v in enumerate(vals, 2):
                c = ws.Cells(r + 2 + i, j); c.NumberFormat = '@'; c.Value = v; c.HorizontalAlignment = -4108
        box(r + 1, 1, 4, 4)
        r += 7
    ws.Cells(r, 1).Value = '■ 비교용 숫자 (위 표의 최소 · 최대) — 다른 시트 함수가 참조할 칸'; ws.Cells(r, 1).Font.Bold = True
    r += 1
    for g, rows in DATA.items():
        ws.Cells(r, 1).Value = g; ws.Cells(r, 1).Font.Bold = True
        head(r + 1, COLS)
        k = r + 2
        for p, vals in zip(PAR, rows):
            for t, idx in (('최소', 0), ('최대', 1)):
                ws.Cells(k, 1).Value = f'{p} {t}'
                for j, v in enumerate(vals, 2):
                    c = ws.Cells(k, j); c.Value = mm(v)[idx]; c.Interior.Color = GRAY; c.HorizontalAlignment = -4108
                    if p.startswith('단위'):
                        c.NumberFormat = '0.0'
                k += 1
        box(r + 1, 1, 7, 4)
        r = k + 1
    ws.Range('A' + str(r)).Value = '※ 풍화토는 본문 표처럼 N≤30 · N＞30 구분 없이 한 범위. 편마암 표토/붕적층 모래는 이번 과업에 없는 지층이지만 본문 표 그대로 둠'
    ws.Columns(1).ColumnWidth = 22
    for j in (2, 3, 4):
        ws.Columns(j).ColumnWidth = 14
    ws.Tab.Color = 0x84B0F4
    ws.PageSetup.Orientation = 1
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    print('시트', [s.Name for s in wb.Worksheets], '| 선택', xl.ActiveWindow.SelectedSheets.Count)
    for rr in range(3, r + 1):
        v = [ws.Cells(rr, c).Text.replace('\n', ' ') for c in range(1, 5)]
        if any(v): print(rr, v)
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
