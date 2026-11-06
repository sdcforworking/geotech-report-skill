"""'관련 문헌자료' 시트 맨 아래에 변형계수 · 포아송비 문헌 표 4개 추가 (본문 ④ 변형계수 및 포아송비 산정의 문헌자료 검토 값 그대로).
본문 표에서 합쳐진 칸(Das 자갈 E·ν, Bowles 자갈·점토 ν, Hunt ν)은 hwp 표 구조를 읽어 칸마다 채움. 위쪽 표 칸 위치는 그대로(다른 시트가 참조).
사용: python lit_e.py 파일.xlsx"""
import os, sys
import win32com.client as w
P = os.path.abspath(sys.argv[1])
BLUE = 0xF2E1D9
DAS = [('느슨한 모래', '10∼24', '0.20∼0.40', '10∼25', '0.20∼0.35'),
       ('보통굳기 모래', '17∼28', '0.25∼0.40', '-', '-'),
       ('조밀한 모래', '35∼55', '0.30∼0.45', '50∼81', '0.30∼0.40'),
       ('느슨한 모래섞인 자갈', '69∼172', '0.15∼0.35', '50∼150', '0.30∼0.40'),
       ('조밀한 모래섞인 자갈', '69∼172', '0.15∼0.35', '100∼200', '0.30∼0.40'),
       ('연약점토', '2∼5', '-', '2∼25', '0.40∼0.50'),
       ('중간점토', '5∼10', '0.20∼0.50', '15∼50', '0.40∼0.50'),
       ('견고점토', '10∼24', '-', '50∼100', '0.40∼0.50')]
HUNT = [('모래', '느슨', '10∼30', '0.20∼0.35'), ('모래', '중간 조밀', '30∼50', '0.20∼0.35'), ('모래', '조밀', '50∼80', '0.20∼0.35'),
        ('자갈', '느슨', '30∼80', '0.30∼0.40'), ('자갈', '중간 조밀', '80∼100', '0.30∼0.40'), ('자갈', '조밀', '100∼120', '0.30∼0.40')]
KDS = (['느슨한 모래', '중간조밀 모래', '조밀한 모래', '실트질 모래', '모래 및 자갈'], ['10∼24', '17∼28', '35∼55', '10∼17', '69∼172'])
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); ws = wb.Worksheets('관련 문헌자료')
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    assert ws.Cells(last, 1).Value == 'ML', (last, ws.Cells(last, 1).Value)

    def table(r0, title, hdr, rows):
        ws.Cells(r0, 1).Value = title; ws.Cells(r0, 1).Font.Bold = True
        for j, h in enumerate(hdr, 1):
            c = ws.Cells(r0 + 1, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = BLUE
            c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
        ws.Rows(r0 + 1).RowHeight = 30
        for i, row in enumerate(rows):
            for j, v in enumerate(row, 1):
                c = ws.Cells(r0 + 2 + i, j); c.NumberFormat = '@'; c.Value = v; c.HorizontalAlignment = -4108
        rng = ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r0 + 1 + len(rows), len(hdr)))
        for b in (7, 8, 9, 10, 11, 12):
            rng.Borders(b).LineStyle = 1; rng.Borders(b).Weight = 2; rng.Borders(b).Color = 0x808080
        return r0 + 2 + len(rows)

    r = last + 3
    ws.Cells(r, 1).Value = '◆ 변형계수 · 포아송비 문헌자료 (본문 ④ 변형계수 및 포아송비 산정)'; ws.Cells(r, 1).Font.Bold = True; ws.Cells(r, 1).Font.Size = 12
    r += 2
    r = table(r, '■ ⑤ Braja M. Das(1984) · Joseph E. Bowles(1996)의 제안',
              ['구분', 'Das 변형계수\n(MPa)', 'Das 포아송비', 'Bowles 변형계수\n(MPa)', 'Bowles 포아송비'], DAS)
    ws.Cells(r, 1).Value = ('※ 본문 표 병합 칸: Das 69∼172 · 0.15∼0.35는 느슨·조밀한 모래섞인 자갈 두 줄, Bowles 0.30∼0.40은 자갈 두 줄, '
                            'Bowles 0.40∼0.50은 점토 세 줄에 걸침 → 칸마다 채움')
    r += 2
    r = table(r, '■ ⑥ Hunt의 제안 (Geotechnical Engineering Analysis & Evaluation, p134)',
              ['토층', '상태', '변형계수\n(MPa)', '포아송비'], HUNT)
    ws.Cells(r, 1).Value = '※ 본문 표에서 포아송비는 모래 · 자갈 각각 세 상태에 걸친 병합 칸 → 칸마다 채움'
    r += 2
    r = table(r, '■ ⑦ 구조물 기초 설계기준 (건설교통부)', ['흙의 종류'] + KDS[0], [tuple(['변형계수(MPa)'] + KDS[1])])
    r += 1
    r = table(r, '■ ⑧ 서울시 지반분류 (2006) — 변형계수 · 포아송비', ['구분', '변형계수\n(MPa)', '포아송비'], [('풍화토', '20∼40', '0.50')])
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    print('추가', last + 3, '∼', r, '| 선택', xl.ActiveWindow.SelectedSheets.Count)
    for rr in range(last + 3, r + 1):
        v = [ws.Cells(rr, c).Text.replace('\n', '') for c in range(1, 7)]
        if any(v): print(rr, v)
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
