"""'산정결과 정리' 시트 — 본문 '구간별 단위중량 및 강도정수 산정결과' 표 형식(행 = 지층 · 정수, 열 = 관련 문헌자료 · 경험식(평균) · 기존 설계사례 ·
현장/실내시험 · 적용 설계정수)으로 다섯 정수를 모음. 앞 네 칸은 다른 시트와 함수로 연결, 적용 설계정수(노란 칸)는 사용자가 정한 값(강도정수)을 넣고 변형계수 · 포아송비는 비움.
사용: python result_sheet.py 파일.xlsx"""
import os, sys
import win32com.client as w

P = os.path.abspath(sys.argv[1])
LAYERS = [('화강암', '붕적층', '모래'), ('화강암', '붕적층', '자갈'), ('화강암', '풍화토', 'N≤30'), ('화강암', '풍화토', 'N＞30'),
          ('편마암', '붕적층', '자갈'), ('편마암', '풍화토', 'N≤30'), ('편마암', '풍화토', 'N＞30')]
PAR = ['단위중량(kN/m³)', '점착력(kPa)', '내부마찰각(˚)', '변형계수(MPa)', '포아송비']
# 사용자가 정한 강도정수 적용값 (대화에서 준 표)
APPLY = {('화강암', '붕적층', '모래'): (18.0, 5.0, 28.0), ('화강암', '붕적층', '자갈'): (19.0, 0, 34.0),
         ('화강암', '풍화토', 'N≤30'): (18.0, 10.0, 28.0), ('화강암', '풍화토', 'N＞30'): (19.0, 20.0, 30.0),
         ('편마암', '붕적층', '자갈'): (19.0, 0, 35.0), ('편마암', '풍화토', 'N≤30'): (18.0, 10.0, 28.0),
         ('편마암', '풍화토', 'N＞30'): (19.0, 20.0, 30.0)}
CASE_ROW0 = {'화강암': 5, '편마암': 14}            # '기존 설계사례 요약' 정수 첫 행
CASE_COL = {'모래': 'B', '자갈': 'C', 'N≤30': 'D', 'N＞30': 'D'}
REASON = {
    '화강암': ['① 단위중량',
              '∙붕적층과 풍화토(N≤30)는 관련 문헌자료 범위 내에 있는 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간보다 조밀한 상태를 고려하여 기존 설계사례 범위 내에서 적용',
              '② 점착력',
              '∙붕적층(모래)은 비소성(NP)이나 실트 등 세립분 혼재를 고려하여 기존 설계사례 범위 내에서 적용',
              '∙붕적층(자갈)은 관련 문헌자료와 기존 설계사례가 모두 0kPa로 점착력을 고려하지 않음',
              '∙풍화토(N≤30)는 관련 문헌자료 범위 내에 있는 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간 적용값과 조밀도 증가를 고려하여 기존 설계사례 범위 내에서 적용',
              '③ 내부마찰각',
              '∙N값 경험식이 관련 문헌자료와 기존 설계사례보다 크게 산정되어 문헌자료와 기존 설계사례를 우선하여 적용',
              '∙붕적층(모래)은 관련 문헌자료 범위 내에서 기존 설계사례의 최대값을 적용',
              '∙붕적층(자갈)은 경험식 최소값이 기존 설계사례 범위 내에 있어 경험식 최소값을 안전측으로 절사하여 적용',
              '∙풍화토(N≤30)는 관련 문헌자료 범위 내에서 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간 적용값을 고려하여 관련 문헌자료와 기존 설계사례 범위 내에서 적용',
              '④ 변형계수 · ⑤ 포아송비 : (적용값 정한 뒤 작성)'],
    '편마암': ['① 단위중량',
              '∙붕적층(자갈)과 풍화토(N≤30)는 관련 문헌자료 범위 내에 있는 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간보다 조밀한 상태를 고려하여 기존 설계사례 범위 내에서 적용',
              '② 점착력',
              '∙붕적층(자갈)은 관련 문헌자료와 기존 설계사례가 모두 0kPa로 점착력을 고려하지 않음',
              '∙풍화토(N≤30)는 관련 문헌자료 범위 내에 있는 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간 적용값과 조밀도 증가를 고려하여 기존 설계사례 범위 내에서 적용',
              '③ 내부마찰각',
              '∙N값 경험식이 관련 문헌자료와 기존 설계사례보다 크게 산정되어 문헌자료와 기존 설계사례를 우선하여 적용',
              '∙붕적층(자갈)은 경험식 최소값이 관련 문헌자료 값과 거의 같아 관련 문헌자료 값을 적용',
              '∙풍화토(N≤30)는 관련 문헌자료 범위 내에서 기존 설계사례의 최소값을 적용',
              '∙풍화토(N＞30)는 N≤30 구간 적용값을 고려하여 관련 문헌자료와 기존 설계사례 범위 내에서 적용',
              '④ 변형계수 · ⑤ 포아송비 : (적용값 정한 뒤 작성)']}
BLUE, YEL, GRN = 0xF2E1D9, 0xCCF2FF, 0xDAEFE2
FMT = ['0.0', '0.0', '0.0', '0', '0.00']

xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(P)
    ws = wb.Worksheets.Add(None, wb.Worksheets('문헌 범위 정리')); ws.Name = '산정결과 정리'
    ws.Cells.Font.Name = '맑은 고딕'; ws.Cells.Font.Size = 10
    ws.Range('A1').Value = '구간별 설계지반정수 산정결과 (본문 표 형식) — 관련 문헌자료 · 경험식 · 기존 설계사례 · 현장시험은 다른 시트와 함수로 연결'
    ws.Range('A1').Font.Size = 13; ws.Range('A1').Font.Bold = True
    ws.Range('A2').Value = ("※ 노란 칸(적용 설계정수)만 직접 넣음 — 강도정수는 정하신 값, 변형계수 · 포아송비는 비워 둠. "
                            "경험식 칸은 '범위(평균)', 평균은 변형계수는 최대값 제외 평균")
    r = 4
    for g in ('화강암', '편마암'):
        ws.Cells(r, 1).Value = f'■ {g}'; ws.Cells(r, 1).Font.Bold = True; ws.Cells(r, 1).Font.Size = 11
        r += 1
        hdr = ['지층', '구분', '정수', '관련 문헌자료', '경험식\n(평균)', '기존\n설계사례', '현장/\n실내시험', '적용\n설계정수', '확인할 점']
        for j, h in enumerate(hdr, 1):
            c = ws.Cells(r, j); c.Value = h; c.Font.Bold = True; c.Interior.Color = YEL if j == 8 else BLUE
            c.HorizontalAlignment = -4108; c.VerticalAlignment = -4108; c.WrapText = True
        ws.Rows(r).RowHeight = 30
        h0 = r; r += 1
        for li, (gg, ly, sub) in enumerate(LAYERS):
            if gg != g:
                continue
            for pi, par in enumerate(PAR):
                ws.Cells(r, 1).Value = ly; ws.Cells(r, 2).Value = sub; ws.Cells(r, 3).Value = par
                lr = 6 + li * 5 + pi                                    # '문헌 범위 정리' 요약 행
                ws.Cells(r, 4).Formula = f"='문헌 범위 정리'!E{lr}"
                pr = 6 + li                                              # 내부마찰각 · 변형계수 시트 행
                if pi == 2:
                    ws.Cells(r, 5).Formula = (f"=IF(ISNUMBER('N값 내부마찰각'!L{pr}),TEXT('N값 내부마찰각'!L{pr},\"0.0\")&\"∼\"&"
                                              f"TEXT('N값 내부마찰각'!M{pr},\"0.0\")&\"(\"&TEXT('N값 내부마찰각'!O{pr},\"0.0\")&\")\",\"-\")")
                elif pi == 3:
                    ws.Cells(r, 5).Formula = (f"=IF(ISNUMBER('N값 변형계수'!N{pr}),TEXT('N값 변형계수'!N{pr},\"0.0\")&\"∼\"&"
                                              f"TEXT('N값 변형계수'!O{pr},\"0.0\")&\"(\"&TEXT('N값 변형계수'!Q{pr},\"0.0\")&\")\",\"-\")")
                else:
                    ws.Cells(r, 5).Value = '-'
                ws.Cells(r, 6).Formula = f"='기존 설계사례 요약'!{CASE_COL[sub]}{CASE_ROW0[g] + pi}"
                if g == '화강암' and sub == 'N＞30' and pi == 3:
                    ws.Cells(r, 7).Value = '79.1∼97.2(89.9)'
                    ws.Cells(r, 9).Value = '공내재하시험 4회 (4.2.2 정리 엑셀 공내재하시험)'
                else:
                    ws.Cells(r, 7).Value = '-'
                if pi < 3:
                    ws.Cells(r, 8).Value = APPLY[(gg, ly, sub)][pi]
                ws.Cells(r, 8).NumberFormat = FMT[pi]; ws.Cells(r, 8).Interior.Color = YEL
                if g == '편마암' and sub == 'N≤30' and pi == 2:
                    ws.Cells(r, 9).Value = '기존 설계사례 29∼30 — 적용 28이 사례 최소보다 낮음 (선정사유 확인)'
                for j in range(4, 9):
                    ws.Cells(r, j).HorizontalAlignment = -4108
                r += 1
        rng = ws.Range(ws.Cells(h0, 1), ws.Cells(r - 1, 9))
        for b in (7, 8, 9, 10, 11, 12):
            rng.Borders(b).LineStyle = 1; rng.Borders(b).Weight = 2; rng.Borders(b).Color = 0x808080
        ws.Range(ws.Cells(h0 + 1, 1), ws.Cells(r - 1, 3)).HorizontalAlignment = -4108
        r += 1
        ws.Cells(r, 1).Value = '선정사유'; ws.Cells(r, 1).Font.Bold = True; r += 1
        for t in REASON[g]:
            ws.Cells(r, 1).Value = t
            if t[0] in '①②③④':
                ws.Cells(r, 1).Font.Bold = True
            r += 1
        r += 2
    for j, wd in enumerate([8, 8, 15, 14, 18, 12, 16, 10, 48], 1):
        ws.Columns(j).ColumnWidth = wd
    ws.Tab.Color = 0x84B0F4
    ws.PageSetup.Orientation = 1
    ws.Activate(); ws.Select(); ws.Range('A1').Select()
    xl.CalculateFull()
    print('선택', xl.ActiveWindow.SelectedSheets.Count, '| 마지막 행', r)
    for rr in range(4, r):
        v = [ws.Cells(rr, c).Text.replace('\n', '') for c in range(1, 10)]
        if any(v[1:]) or v[0].startswith('■'): print(rr, v)
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
