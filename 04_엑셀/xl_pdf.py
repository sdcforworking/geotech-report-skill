import os, sys, fitz
import win32com.client as w
xl = w.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
try:
    for f in sys.argv[1:]:
        b = xl.Workbooks.Open(os.path.abspath(f), 0, True)
        p = os.path.abspath(f[:-5] + '.pdf'); b.ExportAsFixedFormat(0, p); b.Close(False)
        d = fitz.open(p); print(f, '쪽', d.page_count)
        for i in (0, d.page_count - 1):
            d[i].get_pixmap(dpi=80).save(f[:-5] + f'_p{i+1}.png')
finally:
    xl.Quit()
