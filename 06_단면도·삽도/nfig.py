"""n값.ai 양식으로 암종별(화성암·변성암) N값-심도 삽도 4장(붕적층, 풍화대)을 만든다. ai(PDF 호환) + jpg(300dpi).
점 자료: 4.2.2 시추조사결과 엑셀 '암종별 N값 자료' 시트.
사용: python nfig.py 엑셀경로 출력폴더 [폭mm 높이mm]   (기본 79 × 38mm; 양식 n값.ai 는 53 × 41mm)
글자·마커·범례 크기와 그래프 틀 바깥 여백은 양식 그대로 두고, 그래프 틀만 용지 크기에 맞춰 늘이거나 줄인다."""
import sys, os, json
import fitz
import openpyxl

XL, OUT = sys.argv[1], sys.argv[2]
WMM, HMM = (float(sys.argv[3]), float(sys.argv[4])) if len(sys.argv) > 4 else (79.0, 38.0)
here = os.path.dirname(os.path.abspath(__file__))
NAI = os.path.join(here, 'nai')
MK = json.load(open(os.path.join(NAI, 'markers.json')))
FONT_R, FONT_B = 'C:/Windows/Fonts/malgun.ttf', 'C:/Windows/Fonts/malgunbd.ttf'
FR, FB = fitz.Font(fontfile=FONT_R), fitz.Font(fontfile=FONT_B)
W, H = WMM * 72 / 25.4, HMM * 72 / 25.4
X0, Y0 = 22.423, 22.227                                     # 그래프 틀 왼쪽·위 (양식과 같은 여백)
XR, YB = W - 8.504, H - 8.534                               # 그래프 틀 오른쪽·아래 (양식과 같은 여백)
GX = [X0 + i * (XR - X0) / 4 for i in range(5)]             # 눈금 0/30·20/30·40/30·50/20·50/1
K = (GX[1] - X0) / 20                                       # N 1회당
X50, X5020 = X0 + 50 * K, GX[3]                             # 50/30, 50/20 위치
GRIDX, GREY = GX[1:4], (0.4, 0.4, 0.4)
NAVY, RED = (0.083, 0.261, 0.575), (0.902, 0.122, 0.098)


def xpos(n, p):
    """가로축: 0/30∼50/30은 N에 비례, 50/30∼50/20, 50/20∼50/1은 관입량에 비례 (양식 눈금 0/30·20/30·40/30·50/20·50/1)."""
    if n < 50:
        assert p == 30, (n, p)
        return X0 + n * K
    if p >= 20:
        return X50 + (30 - p) / 10 * (X5020 - X50)
    return X5020 + (20 - p) / 19 * (XR - X5020)


def rrect(sh, x0, y0, x1, y1, r=1.598):
    c = r * 0.55
    sh.draw_bezier((x0, y0 + r), (x0, y0 + r - c), (x0 + r - c, y0), (x0 + r, y0))
    sh.draw_line((x0 + r, y0), (x1 - r, y0))
    sh.draw_bezier((x1 - r, y0), (x1 - r + c, y0), (x1, y0 + r - c), (x1, y0 + r))
    sh.draw_line((x1, y0 + r), (x1, y1 - r))
    sh.draw_bezier((x1, y1 - r), (x1, y1 - r + c), (x1 - r + c, y1), (x1 - r, y1))
    sh.draw_line((x1 - r, y1), (x0 + r, y1))
    sh.draw_bezier((x0 + r, y1), (x0 + r - c, y1), (x0, y1 - r + c), (x0, y1 - r))
    sh.draw_line((x0, y1 - r), (x0, y0 + r))


class Fig:
    def __init__(self):
        self.doc = fitz.open()
        self.p = self.doc.new_page(width=W, height=H)
        self.p.insert_font(fontname='MR', fontfile=FONT_R)
        self.p.insert_font(fontname='MB', fontfile=FONT_B)
        self.xref = {}
        self.done = set()                      # 같은 자리에 겹친 시험은 한 번만 그림 (그림자가 겹쳐 진해지지 않게)
        self.pts = []

    def img(self, key, rect, path):
        if key in self.xref:
            self.p.insert_image(fitz.Rect(rect), xref=self.xref[key])
        else:
            self.xref[key] = self.p.insert_image(fitz.Rect(rect), filename=path)

    def text(self, s, x, base, size, bold=False, align='l', rotate=0):
        f = FB if bold else FR
        if align == 'r':
            x -= f.text_length(s, size)
        self.p.insert_text((x, base), s, fontname='MB' if bold else 'MR', fontsize=size, color=(0, 0, 0), rotate=rotate)

    def marker(self, kind, cx, cy, data=True):
        if data:
            key = (kind, round(cx, 2), round(cy, 2))
            if key in self.done:
                return
            self.done.add(key); self.pts.append((cx, cy))
        m = MK[kind]
        a, b, c, d = m['shadow']
        self.img('sh_' + kind, (cx + a, cy + b, cx + c, cy + d), os.path.join(NAI, f'shadow_{kind}.png'))
        sh = self.p.new_shape()
        P = lambda v: (cx + v[1], cy + v[2])
        for it in m['items']:
            if it[0] == 'l':
                sh.draw_line(P(it[1]), P(it[2]))
            elif it[0] == 'c':
                sh.draw_bezier(P(it[1]), P(it[2]), P(it[3]), P(it[4]))
            elif it[0] == 're':
                r = it[1]; sh.draw_rect(fitz.Rect(cx + r[1], cy + r[2], cx + r[3], cy + r[4]))
        sh.finish(fill=m['fill'], color=(1, 1, 1), width=m['w'], closePath=True)
        sh.commit()


def background(f):
    """양식의 그래프 바탕(위 흰색 → 아래 연회색) 그라데이션. 0.5pt 띠를 겹쳐 그림 (벡터)."""
    top, bot = Y0 - 0.227, YB + 0.314
    y = top
    while y < bot:
        t = (y + 0.25 - top) / (bot - top)                  # 양식: 위 21%는 흰색, 아래로 갈수록 239까지
        v = 1.0 if t <= 0.21 else (255 - 16 * (t - 0.21) / 0.79) / 255
        sh = f.p.new_shape(); sh.draw_rect(fitz.Rect(X0 - 0.423, y, XR + 0.268, min(bot, y + 0.55)))
        sh.finish(fill=(v, v, v), color=None, width=0); sh.commit()
        y += 0.5


def frame(f, axmax):
    sh = f.p.new_shape(); sh.draw_rect(fitz.Rect(X0, Y0, XR, YB)); sh.finish(color=(0, 0, 0), width=0.298); sh.commit()
    sh = f.p.new_shape()
    for x in GRIDX:
        sh.draw_line((x, Y0), (x, YB))
    for k in (1, 2):
        y = Y0 + k * (YB - Y0) / 3
        sh.draw_line((X0 + 0.04, y), (XR, y))
    sh.finish(color=GREY, width=0.099, dashes='[0.994 0.497] 0', closePath=False); sh.commit()
    t = 'N값(회/cm)'
    f.text(t, (X0 + XR) / 2 - FB.text_length(t, 7.59) / 2, 9.36, 7.59, bold=True)
    t = '심도(GL.-m)'
    f.text(t, 9.93, (Y0 + YB) / 2 + FB.text_length(t, 7.59) / 2, 7.59, bold=True, rotate=90)
    # 가로 눈금 글자: 양 끝은 틀 안쪽으로 붙이고(양식 위치), 가운데 셋은 눈금 가운데
    labels = ('0/30', '20/30', '40/30', '50/20', '50/1')
    for i, s in enumerate(labels):
        w = FR.text_length(s, 6.64)
        x = X0 - 1.42 if i == 0 else XR + 1.47 - w if i == 4 else GX[i] - w / 2
        f.text(s, x, 19.29, 6.64)
    # 세로 눈금 글자: 양식처럼 맨 위는 아래로, 맨 아래는 위로 조금 당김
    step = axmax / 3
    for k in range(4):
        tick = Y0 + k * (YB - Y0) / 3
        base = Y0 + 4.773 if k == 0 else YB - 0.216 if k == 3 else tick + 2.213
        f.text(f'{step * k:g}', 20.3, base, 6.64, align='r')


# 범례 양식 (상자 왼쪽 위 기준 상대값): 상자 폭·높이, 그림자 여백, 마커 x·첫 줄 y, 글자 x·첫 줄 기준선, 줄 간격, 글자 크기
LEG = {'leg0': dict(w=41.034, h=18.634, sh=(-3.631, -3.632, 3.535, 3.734), mx=6.24, my=5.888, tx=10.0, ty=7.568, step=7.5, size=5.12),
       'leg2': dict(w=29.841, h=18.98, sh=(-3.456, -3.491, 3.703, 3.729), mx=6.224, my=5.699, tx=10.374, ty=7.889, step=8.28, size=5.69)}


def shadow_png(style, nrow):
    """범례 그림자 이미지. 줄 수가 양식(2줄)보다 적으면 가운데 띠를 잘라 높이를 줄인다."""
    src = os.path.join(NAI, style + '.png')
    if nrow == 2:
        return src
    from PIL import Image
    im = Image.open(src); w, h = im.size
    L = LEG[style]
    cut = round(h * L['step'] * (2 - nrow) / (L['h'] + L['sh'][3] - L['sh'][1]))
    top, bot = im.crop((0, 0, w, (h - cut) // 2)), im.crop((0, (h - cut) // 2 + cut, w, h))
    out = Image.new('RGBA', (w, h - cut)); out.paste(top, (0, 0)); out.paste(bot, (0, top.size[1]))
    dst = os.path.join(NAI, f'{style}_{nrow}row.png'); out.save(dst)
    return dst


def legend(f, style, rows, prefer):
    """rows = [(마커, 글자)]. 점을 가리지 않는 모서리에 둔다 (prefer 모서리 먼저: TR, BR, BL, TL)."""
    L = LEG[style]
    w, h = L['w'], L['h'] - L['step'] * (2 - len(rows))
    corners = {'TR': (XR - 6.767 - w, Y0 + 6.705), 'BR': (XR - 6.767 - w, YB - 5.415 - h),
               'BL': (X0 + 5.533, YB - 5.415 - h), 'TL': (X0 + 5.533, Y0 + 6.705)}
    order = [prefer] + [c for c in ('TR', 'BR', 'BL', 'TL') if c != prefer]
    hits = lambda x0, y0: sum(fitz.Rect(x0 - 2.6, y0 - 2.6, x0 + w + 2.6, y0 + h + 2.6).contains(fitz.Point(x, y)) for x, y in f.pts)
    pos = min(order, key=lambda c: (hits(*corners[c]), order.index(c)))
    x0, y0 = corners[pos]
    print('  범례', pos, '가리는 점', hits(x0, y0))
    a, b, c, d = L['sh']
    f.img('leg', (x0 + a, y0 + b, x0 + w + c, y0 + h + d), shadow_png(style, len(rows)))
    sh = f.p.new_shape(); rrect(sh, x0, y0, x0 + w, y0 + h); sh.finish(color=(1, 1, 1), width=2.114, closePath=True); sh.commit()
    sh = f.p.new_shape(); rrect(sh, x0, y0, x0 + w, y0 + h); sh.finish(fill=(1, 1, 1), color=NAVY, width=0.7, closePath=True); sh.commit()
    for i, (kind, label) in enumerate(rows):
        f.marker(kind, x0 + L['mx'], y0 + L['my'] + L['step'] * i, data=False)
        f.text(label, x0 + L['tx'], y0 + L['ty'] + L['step'] * i, L['size'])
    return fitz.Rect(x0 + a, y0 + b, x0 + w + c, y0 + h + d)


def axis_max(deps):
    m = max(deps)
    for s in (1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 20):
        if 3 * s - m >= 0.5:
            return 3 * s


def save(f, name):
    os.makedirs(os.path.join(OUT, 'ai'), exist_ok=True); os.makedirs(os.path.join(OUT, 'jpg'), exist_ok=True)
    f.doc.subset_fonts()
    f.doc.set_metadata({'title': name, 'creator': 'nfig.py'})
    f.doc.save(os.path.join(OUT, 'ai', name + '.ai'), garbage=3, deflate=True)
    pix = f.p.get_pixmap(dpi=300); pix.set_dpi(300, 300)
    pix.save(os.path.join(OUT, 'jpg', name + '.jpg'), jpg_quality=95)
    print('저장', name, '심도축', f.axmax, '시험', f.n, '그린 자리', len(f.pts))


# ── 자료
wb = openpyxl.load_workbook(XL, read_only=True)
ws = wb['암종별 N값 자료']
D = {}
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] in ('화성암', '변성암'):
        n, p = map(int, r[9].split('/'))
        D.setdefault(r[0], []).append(dict(ly=r[4], sub=r[5], dep=float(r[8]), n=n, p=p))
assert sum(len(v) for v in D.values()) == 720

for kind in ('화성암', '변성암'):
    S = D[kind]
    # 1) 붕적층
    pts = [s for s in S if s['ly'] == '붕적층']
    f = Fig(); f.axmax = axis_max([s['dep'] for s in pts]); f.n = len(pts)
    background(f); frame(f, f.axmax)
    Y = lambda dep: Y0 + dep / f.axmax * (YB - Y0)
    subs = [(sub, mk) for sub, mk in (('모래', 'plus'), ('자갈', 'x')) if any(s['sub'] == sub for s in pts)]
    for sub, mk in subs:
        for s in pts:
            if s['sub'] == sub:
                f.marker(mk, xpos(s['n'], s['p']), Y(s['dep']))
    legend(f, 'leg0', [(mk, f'붕적층({sub})') for sub, mk in subs], 'TR')
    save(f, f'N값_{kind}_붕적층')
    # 2) 풍화대
    pts = [s for s in S if s['ly'] in ('풍화토', '풍화암')]
    f = Fig(); f.axmax = axis_max([s['dep'] for s in pts]); f.n = len(pts)
    background(f); frame(f, f.axmax)
    Y = lambda dep: Y0 + dep / f.axmax * (YB - Y0)
    xc = xpos(50, 10)
    sh = f.p.new_shape(); sh.draw_line((xc, Y0 + 0.021), (xc, YB + 0.02)); sh.finish(color=RED, width=0.846, dashes='[2.539 0.846] 0', closePath=False); sh.commit()
    for ly, mk in (('풍화토', 'circle'), ('풍화암', 'square')):
        for s in pts:
            if s['ly'] == ly:
                f.marker(mk, xpos(s['n'], s['p']), Y(s['dep']))
    leg = legend(f, 'leg2', [('circle', '풍화토'), ('square', '풍화암')], 'BL')
    # 풍화암기준 상자: 양식처럼 기준선 위쪽 가운데에 두되, 점을 가리면 점 없는 곳으로 내림/왼쪽으로 옮김
    BW, BH = 140.972 - 112.209, 42.078 - 27.22
    def hits(x0, y0):
        r = fitz.Rect(x0 - 2.6, y0 - 2.6, x0 + BW + 2.6, y0 + BH + 2.6)
        return sum(r.contains(fitz.Point(x, y)) for x, y in f.pts) + (100 if r.intersects(leg) else 0)
    cands = []
    for y0 in [27.22 + 2 * k for k in range(int((YB - 3 - BH - 27.22) / 2) + 1)]:
        cands.append((min(xc - BW / 2, XR - 0.8 - BW), y0))          # 기준선 가운데
        cands.append((xc - 2.0 - BW, y0))                             # 기준선 왼쪽
    bx, by = min(cands, key=lambda c: (hits(*c), c[1], -c[0]))
    print('  기준 상자', round(bx, 1), round(by, 1), '가리는 점', hits(bx, by))
    dx, dy = bx - 112.209, by - 27.22
    f.img('crit', (110.5 + dx, 25.4 + dy, 142.9 + dx, 43.9 + dy), os.path.join(NAI, 'crit.png'))
    sh = f.p.new_shape(); rrect(sh, bx, by, bx + BW, by + BH, r=1.127)
    sh.finish(fill=(1, 1, 1), color=RED, width=0.5, closePath=True); sh.commit()
    f.text('풍화암기준', 113.56 + dx, 33.16 + dy, 5.66)
    f.text('50/10', 119.97 + dx, 40.36 + dy, 5.66)
    save(f, f'N값_{kind}_풍화대')
