import os, win32com.client as w
rgb = lambda r, g, b: r + g * 256 + b * 65536
REPORT = {'산정 틀': '5.4.1 설계지반정수 (적용값 열 U∼Y)', '대표 N값': '5.3.1 ① 암종별 지층별 대표N값',
          'N값 경험식': '5.3.1 ② N값을 이용한 내부마찰각 · ④ 변형계수 경험식 표', '토사층 정수': '5.3.1 ②·④ 산정결과 표와 선정사유',
          '기존 설계사례': '5.2 유사과업 · 유사지반 설계사례 표 (사용 O만)', '설계사례 요약': '5.2 기존 설계사례 요약 표'}
DATA = ['자료_암종', '자료_SPT', '자료_공내재하', '자료_실내시험', '자료 현황']
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(os.path.abspath('tab_in.xlsx'))
    for ws in wb.Worksheets:
        if ws.Name in REPORT:
            ws.Tab.Color = rgb(255, 153, 0)        # 주황: 보고서에 들어갈 시트
        elif ws.Name in DATA:
            ws.Tab.Color = rgb(155, 194, 230)      # 하늘색: 원자료
        else:
            ws.Tab.Color = rgb(191, 191, 191)      # 회색: 참고 · 검토용
    wo = wb.Worksheets('작업 순서')
    r = wo.UsedRange.Row + wo.UsedRange.Rows.Count + 1
    wo.Cells(r, 1).Value = '■ 시트 탭 색: 주황 = 보고서에 들어갈 시트 / 하늘색 = 원자료 / 회색 = 참고 · 검토용'
    wo.Cells(r, 1).Font.Bold = True
    r += 1
    for c, t in enumerate(['시트', '본문 위치'], 1):
        x = wo.Cells(r, c + 1); x.Value = t; x.Font.Bold = True; x.Interior.Color = rgb(242, 242, 242)
        x.Borders.LineStyle = 1
    for name, where in REPORT.items():
        r += 1
        a, b = wo.Cells(r, 2), wo.Cells(r, 3)
        a.Value = name; b.Value = where
        a.Interior.Color = rgb(255, 230, 204)
        for x in (a, b):
            x.Borders.LineStyle = 1
    wb.Worksheets('산정 틀').Activate()
    wb.Save(); wb.Close(False)
finally:
    xl.Quit()
print('ok')
