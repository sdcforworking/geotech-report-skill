"""종단면도 그림칸 표: 크기 고정만 켠다 (칸·내용은 손대지 않음)."""
import sys, re, win32com.client as w
sys.path.insert(0, '.')
from lay import layout
SRC, DST = sys.argv[1], sys.argv[2]
L = layout(SRC)
a = next(i for i, q in enumerate(L) if q['text'].startswith('(1) 철탑'))
b = next(i for i, q in enumerate(L) if q['text'].startswith('4.2.3 지하수위'))
want, fig = set(), False
for i in range(a, b):
    q = L[i]
    if q['text'].startswith('지층종단면도'): fig = True; continue
    if fig and q['table'] and not q['text']: want.add(i)
    else: fig = False
h = w.gencache.EnsureDispatch('HWPFrame.HwpObject'); h.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
try:
    h.XHwpWindows.Item(0).Visible = False
    h.SetMessageBoxMode(0x2FFF1)
    h.Open(SRC, 'HWP', ''); h.EditMode = 1
    n = done = 0; c = h.HeadCtrl
    while c:
        if c.CtrlID == 'tbl' and c.GetAnchorPos(0).Item('Para') in want:
            n += 1
            pr = c.Properties
            if not pr.Item('ProtectSize'):
                pr.SetItem('ProtectSize', 1); c.Properties = pr; done += 1
        c = c.Next
    print('그림칸 표', len(want), '찾음', n, '새로 고정', done)
    h.SaveAs(DST, 'HWP', '')
finally:
    h.Clear(1); h.Quit()
