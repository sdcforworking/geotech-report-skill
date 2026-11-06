"""암종별 N값 모음 엑셀에 '관련 문헌자료' 시트 추가 — 사용자가 준 본문 '문헌자료 검토' 표 4개(도로설계요령 2009 · 철도설계기준 2011 ·
서울시 지반분류 2006 · Hunt 1987)를 값 그대로. 병합 없이 원문에서 합쳐진 칸(종류·통일분류·재료명)은 줄마다 채움.
철도설계기준 점성토 단위중량은 원문 표에서 상태 구간이 합쳐져 있어 어느 상태에 해당하는지 붙여넣기만으로는 알 수 없음 → 원문 순서대로 따로 적음.
사용: python lit_sheet.py 파일.xlsx"""
import os, sys
import win32com.client as w

P = os.path.abspath(sys.argv[1])
ROAD = [('자갈', '밀실한 것, 입도가 좋은 것', 20.0, '0', '40', 'GW, GP'),
        ('자갈', '밀실하지 않는 것, 입도가 나쁜 것', 18.0, '0', '35', 'GW, GP'),
        ('자갈섞인 모래', '밀실한 것', 21.0, '0', '40', 'GW, GP'),
        ('자갈섞인 모래', '밀실하지 않는 것', 19.0, '0', '35', 'GW, GP'),
        ('모래', '밀실한 것', 20.0, '0', '35', 'SW, SP'),
        ('모래', '밀실하지 않는 것', 18.0, '0', '30', 'SW, SP'),
        ('사질토', '밀실한 것', 19.0, '30 이하', '30', 'SM, SC'),
        ('사질토', '밀실하지 않는 것', 17.0, '0', '25', 'SM, SC'),
        ('점성토', '굳은 것', 18.0, '50 이하', '25', 'ML, CL'),
        ('점성토', '약간 무른 것', 17.0, '30 이하', '20', 'ML, CL'),
        ('점성토', '무른 것', 17.0, '15 이하', '20', 'ML, CL'),
        ('점토 및 실트', '굳은 것', 17.0, '50 이하', '20', 'CH, MH, ML'),
        ('점토 및 실트', '약간 무른 것', 16.0, '30 이하', '15', 'CH, MH, ML'),
        ('점토 및 실트', '무른 것', 14.0, '15 이하', '10', 'CH, MH, ML')]
RAIL_S = [('사질토', '매우 느슨함', '4 이하', '11.0∼16.0'), ('사질토', '느슨함', '4∼10', '14.0∼18.0'),
          ('사질토', '중간상태', '10∼30', '17.0∼20.0'), ('사질토', '조밀함', '30∼50', '17.0∼22.0'),
          ('사질토', '매우조밀함', '50 이상', '20.0∼23.0')]
RAIL_C = [('점성토', '매우연약함', '2 이하'), ('점성토', '연약함', '2∼4'), ('점성토', '중간상태', '4∼8'),
          ('점성토', '견고함', '8∼15'), ('점성토', '매우견고함', '15∼30'), ('점성토', '고결', '30 이상')]
HUNT = []
for code, desc, rows in [('GW', 'Well-graded gravels, gravel-sand mixtures', [('Dense', 75, '90', 22.1, 0.22, 40), ('Medium Dense', 50, '55', 20.8, 0.28, 36), ('Loose', 25, '<28', 19.7, 0.36, 32)]),
                         ('GP', 'Poorly graded gravels, gravel sand mixtures', [('Dense', 75, '70', 20.4, 0.33, 38), ('Medium Dense', 50, '50', 19.2, 0.39, 35), ('Loose', 25, '<20', 18.3, 0.47, 32)]),
                         ('SW', 'Well-graded sands, gravely sands', [('Dense', 72, '65', 18.9, 0.43, 37), ('Medium Dense', 50, '35', 17.9, 0.49, 34), ('Loose', 25, '<15', 17.0, 0.57, 30)]),
                         ('SP', 'Poorly graded sands, gravely sands', [('Dense', 75, '50', 17.6, 0.52, 36), ('Medium Dense', 50, '30', 16.7, 0.60, 33), ('Loose', 25, '<10', 15.9, 0.65, 29)]),
                         ('SM', 'Silty sands', [('Dense', 75, '45', 16.5, 0.62, 35), ('Medium Dense', 50, '25', 15.5, 0.74, 32), ('Loose', 25, '<8', 14.9, 0.80, 29)]),
                         ('ML', 'Inorganic silts, fine sands', [('Dense', 75, '35', 14.9, 0.80, 33), ('Medium Dense', 50, '20', 14.1, 0.90, 31), ('Loose', 25, '<4', 13.5, 1.00, 27)])]:
    for r in rows:
        HUNT.append((code, desc) + r)

BLUE, GRAY = 0xF2E1D9, 0xEFEFEF
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    ws = wb.Worksheets.Add(None, wb.Worksheets('기존 설계사례 요약')); ws.Name = '관련 문헌자료'
    ws.Cells.Font.Name = '맑은 고딕'; ws.Cells.Font.Size = 10
    ws.Range('A1').Value = '관련 문헌자료 (본문 5.3.1 문헌자료 검토 표 값 그대로)'
    ws.Range('A1').Font.Size = 13; ws.Range('A1').Font.Bold = True
    ws.Range('A2').Value = '※ 원문에서 합쳐진 칸(종류 · 통일분류 · 재료명)은 병합하지 않고 줄마다 채움. 숫자 칸은 숫자로, "30 이하" · "<28" 같은 칸은 글자로'

    def table(r0, title, hdr, rows, text_cols=()):
        ws.Cells(r0, 1).Value = title; ws.Cells(r0, 1).Font.Bold = True
        for j, h in enumerate(hdr, 1):
            c = ws.Cells(r0 + 1, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = BLUE
            c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
        ws.Rows(r0 + 1).RowHeight = 30
        for i, row in enumerate(rows):
            for j, v in enumerate(row, 1):
                c = ws.Cells(r0 + 2 + i, j)
                if j in text_cols:
                    c.NumberFormat = '@'
                c.Value = v; c.HorizontalAlignment = -4108
        rng = ws.Range(ws.Cells(r0 + 1, 1), ws.Cells(r0 + 1 + len(rows), len(hdr)))
        for b in (7, 8, 9, 10, 11, 12):
            rng.Borders(b).LineStyle = 1; rng.Borders(b).Weight = 2; rng.Borders(b).Color = 0x808080
        return r0 + 2 + len(rows) + 1

    r = 4
    r = table(r, '■ ① 자연지반의 개략적인 지반정수 (도로설계 요령, 2009)',
              ['종류', '재료의 상태', '단위중량\n(kN/m³)', '점착력\n(kPa)', '내부마찰각\n(°)', '통일분류'], ROAD, text_cols=(4, 5))
    for i in range(len(ROAD)):
        ws.Cells(r - 1 - len(ROAD) + i, 3).NumberFormat = '0.0'
    r = table(r, '■ ② 흙의 단위중량 (철도설계기준, 2011) — 사질토',
              ['구분', '상태', 'N값', '단위중량\n(kN/m³)'], RAIL_S, text_cols=(3, 4))
    r = table(r, '■ ② 흙의 단위중량 (철도설계기준, 2011) — 점성토',
              ['구분', '상태', 'N값'], RAIL_C, text_cols=(3,))
    ws.Cells(r - 1, 1).Value = ('※ 점성토 단위중량은 원문 표에 11.0∼19.0 / 17.0∼20.0 / 17.0∼22.0 세 값이 상태 구간을 합친 칸으로 들어 있음 — '
                                '어느 상태 구간인지는 원문 표로 확인 필요 (이번 과업 토사층은 모두 사질토라 쓰지 않음)')
    r += 1
    r = table(r, '■ ③ 서울시 지반분류 (2006)', ['구분', '단위중량\n(kN/m³)', '점착력\n(kPa)', '내부마찰각\n(°)'],
              [('풍화토', '17.0∼20.0', '0∼100', '25∼30')], text_cols=(2, 3, 4))
    r = table(r, '■ ④ Geotechnical Engineering Investigation (Roy E. Hunt, 1987)',
              ['Material', '설명', 'Compactness', 'DR (%)', 'N', 'γd\n(kN/m³)', 'Void ratio e', 'φ (°)'], HUNT, text_cols=(5,))
    for i in range(len(HUNT)):
        ws.Cells(r - 1 - len(HUNT) + i, 6).NumberFormat = '0.0'
        ws.Cells(r - 1 - len(HUNT) + i, 7).NumberFormat = '0.00'
    for j, wd in enumerate([14, 30, 13, 10, 10, 10, 11, 9], 1):
        ws.Columns(j).ColumnWidth = wd
    ws.Range('B:B').WrapText = False
    ws.Tab.Color = 0x84B0F4
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    print('시트', [s.Name for s in wb.Worksheets], '| 선택', xl.ActiveWindow.SelectedSheets.Count, '| 마지막 행', r)
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
