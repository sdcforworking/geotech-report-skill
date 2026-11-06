"""그림칸 표에 칸맞춤 단면도 jpg 넣기. 4공 구간은 위 표 A-B, A-D / 아래 표 B-C, D-C (기존 배치 순서).
사용: python putfig.py 입력.hwp 출력.hwp [철탑 ...]"""
import sys, os, re, json, shutil
import win32com.client as w
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lay import layout

SRC, DST = sys.argv[1], sys.argv[2]
ONLY = sys.argv[3:]
FIG = 'C:/Users/심대천/Desktop/동해안-신가평 보고서/1. 최종 정리/06_지층단면도/칸맞춤'
TMP = r'C:\Users\Public\hwp_fig'
os.makedirs(TMP, exist_ok=True)
sizes = {r[0]: (r[3], r[4]) for r in json.load(open(FIG + '/sizes.json', encoding='utf-8'))}
cm = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cellmap.json'), encoding='utf-8'))

L = layout(SRC)
a = next(i for i, q in enumerate(L) if q['text'].startswith('(1) 철탑'))
b = next(i for i, q in enumerate(L) if q['text'].startswith('4.2.3 지하수위'))
figs, tw, fig = {}, None, False
for i in range(a, b):
    q = L[i]
    m = re.match(r'\(\d+\) 철탑 (\S+) 구간', q['text'])
    if m:
        tw = m.group(1)
    if q['text'].startswith('지층종단면도'):
        fig = True; continue
    if fig and q['table'] and not q['text']:
        figs.setdefault(tw, []).append((i, q['rc'][1]))
    else:
        fig = False

h = w.gencache.EnsureDispatch('HWPFrame.HwpObject'); h.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
run = lambda a_, n=1: [h.Run(a_) for _ in range(n)]
n_put = 0
try:
    h.XHwpWindows.Item(0).Visible = False
    h.SetMessageBoxMode(0x2FFF1)
    h.Open(SRC, 'HWP', ''); h.EditMode = 1
    for tw, tbls in figs.items():
        if ONLY and tw not in ONLY:
            continue
        names = sorted(nm for nm, n in cm[tw]['secs'])
        cells = [(p, c) for p, ncol in tbls for c in range(ncol)]
        assert len(cells) == len(names), (tw, cells, names)
        for (p, c), nm in zip(cells, names):
            src = f'{FIG}/jpg/지층단면도_{nm}.jpg'
            tmp = os.path.join(TMP, f'f{n_put:03d}.jpg')
            shutil.copyfile(src, tmp)
            h.SetPos(0, p, 0); run('MoveRight')
            if '(A1)' not in h.KeyIndicator()[-1]:
                raise RuntimeError(f'표 진입 실패 {tw} {p}')
            if c:
                run('TableRightCell', c)
            run('MoveListBegin'); run('MoveSelListEnd'); run('Delete'); run('Cancel')
            run('ParagraphShapeAlignCenter')
            bw, bh = sizes[nm]
            ctrl = h.InsertPicture(tmp, True, 1, False, False, 0, bw, bh)
            if ctrl is None:
                raise RuntimeError(f'그림 넣기 실패 {nm}')
            n_put += 1
        run('Cancel'); run('MoveDocBegin')
    print('넣은 그림', n_put, flush=True)
    h.SaveAs(DST, 'HWP', '')
finally:
    h.Clear(1); h.Quit()
