"""보고서 46회 시험의 원본 USCS 시트(Classification of a Soil)를 보고서 순서대로 한 엑셀에 모음.
원본 xls → 임시 xlsx(깨진 이름 정의 삭제) → 시트 복사 → 원본 연결 끊어 값·차트 고정 → 원본과 값·입도곡선 점 대조.
사용: python lab_uscs.py 출력.xlsx"""
import os, sys, json, glob, shutil
import win32com.client as w
import openpyxl
SRC = r'C:\Users\심대천\Desktop\동해안-신가평 보고서\2. 원본자료(차수별)\04_실내시험\SPT총괄표(xls)'
OUT = os.path.abspath(sys.argv[1]); TMP = os.path.abspath('labtmp'); shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
M = json.load(open('lab_map.json', encoding='utf-8'))
T = openpyxl.load_workbook(glob.glob(r'C:\Users\심대천\Desktop\동해안-신가평 보고서\1. 최종 정리\03_실내시험\*.xlsx')[0], data_only=True)['실내시험 총괄표']
ORDER = [(r[2], f'{r[2]}|{float(r[3]):g}') for r in T.iter_rows(min_row=9, values_only=True) if r[2] and str(r[2]).startswith('No.')]
CELLS = ['C4', 'L4', 'C6', 'J6', 'E11', 'E14', 'E17', 'E20', 'E23', 'E26', 'E27', 'K11', 'K14', 'K17', 'K20', 'K23', 'K26']
def snap(s):
    v = {c: s.Range(c).Text for c in CELLS}
    se = s.ChartObjects(1).Chart.SeriesCollection(6)
    v['curve'] = (list(se.XValues), list(se.Values))
    return v
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False; xl.AskToUpdateLinks = False
books, info, bad = {}, [], []
try:
    wb = xl.Workbooks.Add(); wb.Worksheets(1).Name = '_빈시트'
    for hole, k in ORDER:
        fn, sn, _ = M['uscs'][k][0]
        if fn not in books:
            b = xl.Workbooks.Open(os.path.join(SRC, fn), ReadOnly=True, UpdateLinks=0)
            x = os.path.join(TMP, f'{len(books):02d}.xlsx'); b.SaveAs(x, 51); b.Close(False)
            b = xl.Workbooks.Open(x, UpdateLinks=0)
            for n in list(b.Names):
                r = n.RefersTo
                if '#REF' in r or chr(92) in r or '[' in r: n.Delete()
            books[fn] = b
        src = books[fn].Worksheets(sn)
        before = snap(src)
        src.Copy(None, wb.Worksheets(wb.Worksheets.Count))
        ws = wb.Worksheets(wb.Worksheets.Count); ws.Name = hole
        ws.PageSetup.PrintArea = '$A$1:$N$53'
        # 인쇄 범위 밖(O열) 보조 차트: 원본에서도 계열이 깨져 있어 외부 연결로 남으므로 지움
        for co in list(ws.ChartObjects()):
            if co.TopLeftCell.Column > 14:
                co.Delete()
        info.append(dict(hole=hole, file=fn, sheet=sn, before=before))
    wb.Worksheets('_빈시트').Delete()
    for l in (wb.LinkSources(1) or ()):
        wb.BreakLink(l, 1)
    left = wb.LinkSources(1)
    for d in info:
        after = snap(wb.Worksheets(d['hole']))
        if after != d['before']:
            bad.append((d['hole'], {c: (d['before'][c], after[c]) for c in after if after[c] != d['before'][c]}))
    wb.Worksheets(1).Activate()
    wb.SaveAs(OUT, 51); wb.Close(False)
    for b in books.values():
        b.Close(False)
finally:
    xl.Quit()
print('시트', len(info), '| 남은 외부 연결', left, '| 원본과 다른 시트', len(bad))
for x in bad[:5]: print('  ', x)
json.dump([(d['hole'], d['file'], d['sheet']) for d in info], open('lab_uscs_info.json', 'w', encoding='utf-8'), ensure_ascii=False)
