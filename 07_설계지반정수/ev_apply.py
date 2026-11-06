"""'산정결과 정리' 시트: 변형계수 · 포아송비 적용 설계정수(사용자 확정값)를 노란 칸에 넣고, 선정사유 자리표시 줄을 사용자 확정 문구로 바꿈(줄 끼워 넣기). 시트는 지우지 않음.
사용: python ev_apply.py 파일.xlsx"""
import os, sys
import win32com.client as w
P = os.path.abspath(sys.argv[1])
VAL = {('화강암', '붕적층', '모래'): (15, 0.35), ('화강암', '붕적층', '자갈'): (29, 0.35), ('화강암', '풍화토', 'N≤30'): (16, 0.34),
       ('화강암', '풍화토', 'N＞30'): (34, 0.34), ('편마암', '붕적층', '자갈'): (30, 0.35), ('편마암', '풍화토', 'N≤30'): (16, 0.34),
       ('편마암', '풍화토', 'N＞30'): (34, 0.34)}
NU = '∙관련 문헌자료는 편차가 크게 제시되어 안전측 설계를 위하여 토질특성이 유사한 기존 설계사례의 최대값을 적용'
TXT = {'화강암': ['[변형계수 및 포아송비]', '① 변형계수',
                 '∙경험식의 평균값(최대값 제외)을 기준으로 관련 문헌자료와 기존 설계사례를 비교하여 안전측으로 적용',
                 '∙붕적층(모래)은 경험식의 평균값이 기존 설계사례 범위 내에 있으나 안전측으로 기존 설계사례의 최소값을 적용',
                 '∙붕적층(자갈)과 풍화토(N≤30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례보다 작게 산정되어 안전측으로 경험식의 평균값을 적용',
                 '∙풍화토(N＞30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례 범위 내에 있고 공내재하시험 결과보다 작아 경험식의 평균값을 적용',
                 '② 포아송비', NU],
       '편마암': ['[변형계수 및 포아송비]', '① 변형계수',
                 '∙경험식의 평균값(최대값 제외)을 기준으로 관련 문헌자료와 기존 설계사례를 비교하여 안전측으로 적용',
                 '∙붕적층(자갈)은 경험식의 평균값이 관련 문헌자료보다 작고 기존 설계사례와 비슷하여 안전측으로 기존 설계사례 값을 적용',
                 '∙풍화토(N≤30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례보다 작게 산정되어 안전측으로 경험식의 평균값을 적용',
                 '∙풍화토(N＞30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례 범위 내에 있어 경험식의 평균값을 적용',
                 '② 포아송비', NU]}
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P); ws = wb.Worksheets('산정결과 정리')
    g, n = None, 0
    for r in range(1, 120):
        a = ws.Cells(r, 1).Value
        if isinstance(a, str) and a.startswith('■ '):
            g = a[2:]
        key = (g, a, ws.Cells(r, 2).Value)
        par = ws.Cells(r, 3).Value
        if key in VAL and par in ('변형계수(MPa)', '포아송비'):
            assert ws.Cells(r, 8).Value is None, (r, ws.Cells(r, 8).Value)
            ws.Cells(r, 8).Value = VAL[key][0 if par.startswith('변형') else 1]; n += 1
    assert n == 14, n
    # 자리표시 줄 바꾸기 (아래 것부터 — 줄을 끼우면 위치가 밀리므로)
    holders = [r for r in range(1, 150) if str(ws.Cells(r, 1).Value).startswith('④ 변형계수 · ⑤ 포아송비')]
    assert len(holders) == 2, holders
    for r, gg in zip(reversed(holders), ('편마암', '화강암')):
        lines = TXT[gg]
        ws.Rows(f'{r + 1}:{r + len(lines) - 1}').Insert()
        for i, t in enumerate(lines):
            c = ws.Cells(r + i, 1); c.Value = t; c.Font.Bold = t[0] in '①②[' ; c.Interior.Pattern = -4142
    ws.Range('A2').Value = ("※ 노란 칸(적용 설계정수)은 사용자가 정한 값. 관련 문헌자료 · 경험식 · 기존 설계사례 · 현장시험 칸은 다른 시트와 함수로 연결. "
                            "경험식 칸은 '범위(평균)', 변형계수 평균은 최대값 제외 평균")
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    xl.CalculateFull()
    print('적용값 넣은 칸', n, '| 선택', xl.ActiveWindow.SelectedSheets.Count)
    for r in range(1, 110):
        v = [ws.Cells(r, c).Text.replace('\n', '') for c in (1, 2, 3, 8)]
        if v[2] in ('변형계수(MPa)', '포아송비') or v[0].startswith(('■', '선정', '[', '①', '②', '③', '∙')): print(r, v[0][:70], v[1], v[2], v[3])
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
