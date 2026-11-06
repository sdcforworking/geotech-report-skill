"""'산정결과 정리' 풍화토(N＞30) 변형계수 적용값을 사용자가 정한 값(화강암 80, 편마암 60)으로 바꾸고
[변형계수 및 포아송비] 선정사유의 풍화토(N＞30) 줄을 본문에 넣을 문장으로 바꿈. 바꾸기 전 값이 예상과 다르면 멈춤.
사용: python e_n30_apply.py 파일.xlsx"""
import os, sys
import win32com.client as w

sys.stdout.reconfigure(encoding='utf-8')
P = os.path.abspath(sys.argv[1])
EDITS = [  # (칸, 지금 값 확인용, 새 값)
    ('H24', 34, 80),
    ('A47', '∙풍화토(N＞30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례 범위 내에 있고',
     '∙풍화토(N＞30)는 공내재하시험 결과가 경험식의 평균값, 관련 문헌자료 및 기존 설계사례보다 크게 나타나 '
     '현장시험 결과를 우선하되, 안전측으로 시험결과의 최소값 수준을 적용'),
    ('H67', 34, 60),
    ('A88', '∙풍화토(N＞30)는 경험식의 평균값이 관련 문헌자료와 기존 설계사례 범위 내에 있어',
     '∙풍화토(N＞30)는 공내재하시험 결과가 없어, 화강암 풍화토(N＞30)의 공내재하시험 결과보다 작은 범위에서 '
     '기존 설계사례의 최대값을 적용'),
]

xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    ws = wb.Worksheets('산정결과 정리')
    for addr, old, new in EDITS:
        cur = ws.Range(addr).Value
        ok = (cur == old) if isinstance(old, (int, float)) else str(cur).startswith(old)
        if not ok:
            raise SystemExit(f'{addr} 지금 값이 예상과 다름: {cur!r}')
    for addr, old, new in EDITS:
        ws.Range(addr).Value = new
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    xl.CalculateFull()
    print('선택 탭', xl.ActiveWindow.SelectedSheets.Count)
    for addr, _, _ in EDITS:
        print(addr, repr(ws.Range(addr).Value))
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
