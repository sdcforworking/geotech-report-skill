"""'문헌 범위 정리' 시트 — 본문 '구간별 단위중량 및 강도정수 산정결과' 표의 '관련 문헌자료' 칸에 넣을 범위를 정리.
아래 후보 목록: 지층마다 해당하는 문헌 항목(통일분류 · 대표 N값 상대밀도로 고름)을 '관련 문헌자료' 시트 칸에서 함수로 끌어오고, 사용(O/X)을 고름.
위 요약: 본문 표 배치(행 = 지층 · 정수)로, 사용 O인 항목의 최소∼최대와 근거를 함수로 보여 줌. 엑셀(COM)로 작성, 병합 없음, 탭 하나만 선택.
사용: python lit_range.py 파일.xlsx"""
import os, sys
import win32com.client as w

P = os.path.abspath(sys.argv[1])
L = "'관련 문헌자료'!"
ROAD, RAIL, SEOUL, HUNT = '도로설계요령(2009)', '철도설계기준(2011)', '서울시 지반분류(2006)', 'Hunt(1987)'
# (문헌, 분류, 참조칸, 값 종류, 기본 사용)  값 종류: num 숫자 · txt 숫자글자 · le "30 이하" · rng "a∼b"
SAND = {'단위중량': [(ROAD, '사질토(SM·SC) 밀실하지 않은 것', 'C13', 'num', 'O'), (ROAD, '사질토(SM·SC) 밀실한 것', 'C12', 'num', 'X'),
                  (RAIL, '사질토 중간상태 (N 10∼30)', 'E24', 'rng', 'O'), (HUNT, 'SM Medium Dense (N 25) γd', 'F50', 'num', 'O'),
                  (HUNT, 'SM Loose (N＜8) γd', 'F51', 'num', 'X')],
        '점착력': [(ROAD, '사질토(SM·SC) 밀실하지 않은 것', 'D13', 'txt', 'O'), (ROAD, '사질토(SM·SC) 밀실한 것', 'D12', 'le', 'X')],
        '내부마찰각': [(ROAD, '사질토(SM·SC) 밀실하지 않은 것', 'E13', 'txt', 'O'), (ROAD, '사질토(SM·SC) 밀실한 것', 'E12', 'txt', 'X'),
                   (HUNT, 'SM Medium Dense (N 25)', 'H50', 'num', 'O'), (HUNT, 'SM Loose (N＜8)', 'H51', 'num', 'X')]}
GRAV = {'단위중량': [(ROAD, '자갈 밀실하지 않은 것', 'C7', 'num', 'O'), (ROAD, '자갈 밀실한 것', 'C6', 'num', 'X'),
                  (RAIL, '사질토 중간상태 (N 10∼30)', 'E24', 'rng', 'O'), (HUNT, 'GP Medium Dense (N 50) γd', 'F41', 'num', 'O'),
                  (HUNT, 'GP Loose (N＜20) γd', 'F42', 'num', 'X')],
        '점착력': [(ROAD, '자갈 밀실하지 않은 것', 'D7', 'txt', 'O')],
        '내부마찰각': [(ROAD, '자갈 밀실하지 않은 것', 'E7', 'txt', 'O'), (ROAD, '자갈 밀실한 것', 'E6', 'txt', 'X'),
                   (HUNT, 'GP Medium Dense (N 50)', 'H41', 'num', 'O'), (HUNT, 'GP Loose (N＜20)', 'H42', 'num', 'X')]}
WS30 = {'단위중량': [(SEOUL, '풍화토', 'B33', 'rng', 'O'), (RAIL, '사질토 중간상태 (N 10∼30)', 'E24', 'rng', 'X')],
        '점착력': [(SEOUL, '풍화토', 'C33', 'rng', 'O'), (ROAD, '사질토(SM·SC) 밀실하지 않은 것', 'D13', 'txt', 'X')],
        '내부마찰각': [(SEOUL, '풍화토', 'D33', 'rng', 'O'), (HUNT, 'SM Medium Dense (N 25)', 'H50', 'num', 'X')]}
WS50 = {'단위중량': [(SEOUL, '풍화토', 'B33', 'rng', 'O'), (RAIL, '사질토 조밀함 (N 30∼50)', 'F24', 'rng', 'X')],
        '점착력': [(SEOUL, '풍화토', 'C33', 'rng', 'O'), (ROAD, '사질토(SM·SC) 밀실한 것', 'D12', 'le', 'X')],
        '내부마찰각': [(SEOUL, '풍화토', 'D33', 'rng', 'O'), (HUNT, 'SM Dense (N 45)', 'H49', 'num', 'X'),
                   (ROAD, '사질토(SM·SC) 밀실한 것', 'E12', 'txt', 'X')]}
LAYERS = [('화강암', '붕적층', '모래', SAND), ('화강암', '붕적층', '자갈', GRAV), ('화강암', '풍화토', 'N≤30', WS30), ('화강암', '풍화토', 'N＞30', WS50),
          ('편마암', '붕적층', '자갈', GRAV), ('편마암', '풍화토', 'N≤30', WS30), ('편마암', '풍화토', 'N＞30', WS50)]
PAR = [('단위중량', '단위중량(kN/m³)', '0.0'), ('점착력', '점착력(kPa)', '0'), ('내부마찰각', '내부마찰각(˚)', '0')]


def lohi(ref, kind):
    x = L + ref
    if kind == 'num':
        return f'={x}', f'={x}'
    if kind == 'txt':
        return f'=VALUE({x})', f'=VALUE({x})'
    if kind == 'le':
        return '=0', f'=VALUE(LEFT({x},FIND(" ",{x})-1))'
    return f'=VALUE(LEFT({x},FIND("∼",{x})-1))', f'=VALUE(MID({x},FIND("∼",{x})+1,9))'


BLUE, YEL, GRN = 0xF2E1D9, 0xCCF2FF, 0xDAEFE2
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    ws = wb.Worksheets.Add(None, wb.Worksheets('관련 문헌자료')); ws.Name = '문헌 범위 정리'
    ws.Cells.Font.Name = '맑은 고딕'; ws.Cells.Font.Size = 10
    ws.Range('A1').Value = "관련 문헌자료 범위 정리 — 본문 '구간별 단위중량 및 강도정수 산정결과' 표의 '관련 문헌자료' 칸에 넣을 값"
    ws.Range('A1').Font.Size = 13; ws.Range('A1').Font.Bold = True
    ws.Range('A2').Value = ("※ 위 요약표는 모두 함수: 아래 후보 목록에서 사용 'O'인 문헌 값의 최소∼최대. 노란 칸(사용 O/X)만 고르면 됨. "
                            "문헌 값은 '관련 문헌자료' 시트 칸을 그대로 끌어옴")
    ws.Range('A3').Value = ("※ 기본 선택: 대표 N값의 상대밀도(붕적층 23∼27/30 보통조밀 → 도로설계요령 '밀실하지 않은 것' · 철도 '중간상태' · Hunt 'Medium Dense'), "
                            "풍화토는 본문 예시처럼 서울시 지반분류. Hunt 값은 건조단위중량(γd)이라 단위중량 하한 쪽으로 들어감")

    # ── 후보 목록 위치 계산
    n_sum = sum(1 for _ in LAYERS) * 3
    top = 5                                   # 요약표 머리글 행
    cand0 = top + 2 + n_sum + 3              # 후보 목록 머리글 행
    rows = []
    for g, ly, sub, spec in LAYERS:
        for key, _, _ in PAR:
            for src, cls, ref, kind, use in spec[key]:
                rows.append((g, ly, sub, key, src, cls, ref, kind, use))
    c1, c2 = cand0 + 1, cand0 + len(rows)
    K = {k: f'${col}${c1}:${col}${c2}' for k, col in zip(['g', 'ly', 'sub', 'par', 'src', 'cls', 'raw', 'lo', 'hi', 'use'], 'ABCDEFGHIJ')}

    # ── 요약표 (본문 배치)
    hdr = ['암 구분', '지층', '구분', '정수', '관련 문헌자료\n(사용 O 범위)', '최소', '최대', '근거 (사용 O인 문헌)']
    for j, h in enumerate(hdr, 1):
        c = ws.Cells(top, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = GRN if j == 5 else BLUE
        c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
    ws.Rows(top).RowHeight = 30
    r = top + 1
    for g, ly, sub, _ in LAYERS:
        for key, label, fmt in PAR:
            ws.Cells(r, 1).Value = g; ws.Cells(r, 2).Value = ly; ws.Cells(r, 3).Value = sub; ws.Cells(r, 4).Value = label
            crit = f'{K["g"]},A{r},{K["ly"]},B{r},{K["sub"]},C{r},{K["par"]},"{key}",{K["use"]},"O"'
            ws.Cells(r, 6).Formula = f'=IF(COUNTIFS({crit})=0,"",MINIFS({K["lo"]},{crit}))'
            ws.Cells(r, 7).Formula = f'=IF(COUNTIFS({crit})=0,"",MAXIFS({K["hi"]},{crit}))'
            ws.Cells(r, 5).Formula = f'=IF(F{r}="","-",IF(F{r}=G{r},TEXT(F{r},"{fmt}"),TEXT(F{r},"{fmt}")&"∼"&TEXT(G{r},"{fmt}")))'
            cond = f'({K["g"]}=A{r})*({K["ly"]}=B{r})*({K["sub"]}=C{r})*({K["par"]}="{key}")*({K["use"]}="O")'
            ws.Cells(r, 8).FormulaArray = f'=TEXTJOIN(" · ",TRUE,IF({cond},{K["src"]}&" "&{K["cls"]},""))'
            ws.Cells(r, 5).Interior.Color = GRN
            for j in (6, 7):
                ws.Cells(r, j).NumberFormat = fmt
            r += 1
    sr = ws.Range(ws.Cells(top, 1), ws.Cells(r - 1, 8))
    for b in (7, 8, 9, 10, 11, 12):
        sr.Borders(b).LineStyle = 1; sr.Borders(b).Weight = 2; sr.Borders(b).Color = 0x808080
    ws.Range(ws.Cells(top + 1, 1), ws.Cells(r - 1, 7)).HorizontalAlignment = -4108

    # ── 후보 목록
    ws.Cells(cand0 - 1, 1).Value = '■ 후보 목록 — 노란 칸 사용(O/X)을 바꾸면 위 요약표가 따라 바뀜'; ws.Cells(cand0 - 1, 1).Font.Bold = True
    hdr2 = ['암 구분', '지층', '구분', '정수', '문헌', '분류 · 상태', '원문 값', '하한', '상한', '사용\n(O/X)']
    for j, h in enumerate(hdr2, 1):
        c = ws.Cells(cand0, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = YEL if j == 10 else BLUE
        c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
    ws.Rows(cand0).RowHeight = 30
    vals, forms = [], []
    for i, (g, ly, sub, key, src, cls, ref, kind, use) in enumerate(rows):
        rr = c1 + i
        lo, hi = lohi(ref, kind)
        ws.Range(f'A{rr}:F{rr}').Value = (g, ly, sub, key, src, cls)
        ws.Range(f'G{rr}').Formula = f'={L}{ref}&""'
        ws.Range(f'H{rr}').Formula = lo; ws.Range(f'I{rr}').Formula = hi
        ws.Range(f'J{rr}').Value = use
    cr = ws.Range(f'A{cand0}:J{c2}')
    for b in (7, 8, 9, 10, 11, 12):
        cr.Borders(b).LineStyle = 1; cr.Borders(b).Weight = 2; cr.Borders(b).Color = 0x808080
    ws.Range(f'J{c1}:J{c2}').Interior.Color = YEL
    ws.Range(f'A{c1}:J{c2}').HorizontalAlignment = -4108
    ws.Range(f'E{c1}:F{c2}').HorizontalAlignment = -4131
    dv = ws.Range(f'J{c1}:J{c2}').Validation; dv.Delete(); dv.Add(3, 1, 1, 'O,X')
    for j, wd in enumerate([10, 8, 8, 16, 19, 30, 11, 8, 8, 8], 1):
        ws.Columns(j).ColumnWidth = wd
    ws.Columns(8).ColumnWidth = 8
    ws.Range(f'H{top}:H{top + n_sum}').WrapText = False
    ws.Tab.Color = 0x84B0F4
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    xl.CalculateFull()
    print('후보', len(rows), '줄 | 요약', top + 1, '∼', r - 1, '| 후보 머리글', cand0, '| 선택', xl.ActiveWindow.SelectedSheets.Count)
    for rr in range(top + 1, r):
        print([ws.Cells(rr, j).Text for j in (1, 2, 3, 4, 5, 8)])
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
