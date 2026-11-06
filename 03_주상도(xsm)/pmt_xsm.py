"""공내재하시험 결과 요약(통합).xlsx 를 통합 xsm 기준으로 맞춘다 (결과값 D·E는 손대지 않음).
결과시트 머리 칸: J4 G.W.L, G5 Geology, D7 TCR/RQD, D8 N-Value  /  요약 시트: E 지층, F N값 or TCR/RQD.
바뀐 칸에는 메모(이전 값)를 달고, 요약 시트 비고(K)에 적는다. 차트·메모가 있어 엑셀(COM)로 고친다.
사용: python pmt_xsm.py 엑셀경로 [--apply]"""
import sys, os, json
here = os.path.dirname(os.path.abspath(__file__))
X = json.load(open(os.path.join(here, 'xsm_now.json'), encoding='utf-8'))
GW = json.load(open(os.path.join(here, 'gwl.json'), encoding='utf-8'))
F, APPLY = os.path.abspath(sys.argv[1]), '--apply' in sys.argv


def xsm_at(h, dep):
    v = X[h]; top = 0.0
    for i, L in enumerate(v['lay']):
        if top - 1e-6 <= dep < L['bot'] - 1e-6 or i == len(v['lay']) - 1:
            ly = L['ly']; break
        top = L['bot']
    tcr = n = '-'
    if ly == '연암':
        c = [c for c in v['core'] if c[0] - 1e-6 <= dep <= c[1] + 1e-6]
        assert len(c) == 1, (h, dep, c)
        tcr = f'{c[0][2][0]}/{c[0][2][1]}'
    else:
        s = [s for s in v['spt'] if abs(s['dep'] - dep) < 1e-6 and s['n'] > 0]
        n = f"{s[0]['n']}/{int(round(s[0]['p']))}" if s else '-'
    g = GW[h]
    return dict(J4=g if g is not None else '시추심도 이하', G5=ly, D7=tcr, D8=n)


def same(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-9
    return a == b


import win32com.client as w
xl = w.DispatchEx('Excel.Application')          # 사용자 엑셀과 따로 띄움
xl.Visible = False; xl.DisplayAlerts = False; xl.AskToUpdateLinks = False
try:
    wb = xl.Workbooks.Open(F, UpdateLinks=0, ReadOnly=not APPLY)
    xl.Calculation = -4135                      # 수동 계산 (결과값이 다시 계산되지 않게)
    S = wb.Worksheets('공내재하시험 요약')
    rows = {}
    r = 4
    while isinstance(S.Cells(r, 1).Value, float):
        rows[S.Cells(r, 2).Value] = r; r += 1
    before = {h: (S.Cells(rr, 7).Value, S.Cells(rr, 8).Value) for h, rr in rows.items()}
    NAMES = {'J4': 'G.W.L', 'G5': 'Geology', 'D7': 'TCR/RQD', 'D8': 'N-Value'}
    changes = []
    for h, rr in rows.items():
        ws = wb.Worksheets(h)
        dep = ws.Range('D6').Value
        assert abs(dep - S.Cells(rr, 4).Value) < 1e-9
        tgt = xsm_at(h, dep)
        notes = []
        for a, v in tgt.items():
            old = ws.Range(a).Value
            if not same(old, v):
                changes.append((h, NAMES[a], old, v))
                notes.append(f"{NAMES[a]} {'빈칸' if old is None else old} → {v}")
                if APPLY:
                    c = ws.Range(a)
                    c.NumberFormat = '@' if isinstance(v, str) else 'General'
                    c.Value = v
                    if c.Comment is not None:
                        c.Comment.Delete()
                    c.AddComment(f"xsm 기준 수정 (2026-10-27)\n이전: {'빈칸' if old is None else old}")
        if h == 'No.372(A-B)' and ws.Range('G6').Value == '풍화암':     # Geology 값이 Rock Grade 칸에 잘못 들어가 있음
            changes.append((h, 'Rock Grade', '풍화암', '빈칸 (Geology 칸으로 옮김)'))
            notes.append('Rock Grade 풍화암 → 빈칸 (Geology로 옮김)')
            if APPLY:
                ws.Range('G6').Value = None
                ws.Range('G6').AddComment('xsm 기준 수정 (2026-10-27)\n이전: 풍화암 (Geology 칸으로 옮김)')
        # 요약 시트
        sumv = {5: tgt['G5'], 6: tgt['D7'] if tgt['G5'] == '연암' else tgt['D8']}
        for c, v in sumv.items():
            old = S.Cells(rr, c).Value
            if not same(old, v):
                changes.append((h, '요약 ' + ('지층' if c == 5 else 'N값/TCR·RQD'), old, v))
                if APPLY:
                    S.Cells(rr, c).NumberFormat = '@'; S.Cells(rr, c).Value = v
        if notes and APPLY:
            k = S.Cells(rr, 11).Value
            add = 'xsm 기준 수정: ' + ', '.join(notes)
            S.Cells(rr, 11).Value = (k + ' / ' if k else '') + add
    for c in changes:
        print(' | '.join(str(x) for x in c))
    print('바뀌는 칸', len(changes))
    if APPLY:
        after = {h: (S.Cells(rr, 7).Value, S.Cells(rr, 8).Value) for h, rr in rows.items()}
        assert after == before, '결과값이 바뀜'
        wb.Save(); print('저장, 결과값(D·E) 46공 변화 없음')
    wb.Close(False)
finally:
    xl.Quit()
