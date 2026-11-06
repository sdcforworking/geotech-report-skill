"""지층단면도 PDF(.ai 호환)/JPG 생성 — 참고 ai(지층단면도260626(273~280).ai, 274-A 쪽)에서 실측한 양식 사용.

단위: pt. 모든 치수는 참고 ai에서 잰 값.
"""
import math
import fitz

FONT_R = 'C:/Windows/Fonts/malgun.ttf'
FONT_B = 'C:/Windows/Fonts/malgunbd.ttf'
FR, FB = fitz.Font(fontfile=FONT_R), fitz.Font(fontfile=FONT_B)

FILL = {'붕적층': (0.929, 0.886, 0.827), '풍화토': (0.743, 0.715, 0.661),
        '풍화암': (0.665, 0.640, 0.595), '연암': (0.629, 0.574, 0.553)}
ORDER = ['붕적층', '풍화토', '풍화암', '연암']
LABEL = {'연암': '연암'}
SOIL = {'붕적층', '풍화토'}
RED = (0.902, 0.0, 0.071)
DARK = (0.137, 0.094, 0.082)          # 심도 글자색(0x231815)

V = 7.112            # pt / m (세로)
H = 9.582            # pt / m (가로, 공 사이)
HALF = 5.525         # 주상 반폭
SUB = 2.765          # 주상 안 세로선 간격
X1 = 70.47           # 첫 시추공 중심
LW = 0.3


def tw(font, s, size):
    return font.text_length(s, fontsize=size)


class Pen:
    def __init__(self, page):
        self.p = page
        page.insert_font(fontname='MR', fontfile=FONT_R)
        page.insert_font(fontname='MB', fontfile=FONT_B)

    def line(self, a, b, w=LW, color=(0, 0, 0), dashes=None):
        s = self.p.new_shape()
        s.draw_line(a, b)
        s.finish(color=color, width=w, dashes=dashes, lineCap=0)
        s.commit()

    def dashdot(self, a, b, w=0.35, pattern=(5, 1, 1, 1)):
        """1점쇄선을 실제 선분으로 그림 (뷰어마다 점선 해석이 달라도 같게 보이도록)."""
        L = math.dist(a, b)
        ux, uy = (b[0] - a[0]) / L, (b[1] - a[1]) / L
        s = self.p.new_shape()
        t, i = 0.0, 0
        while t < L:
            seg = pattern[i % len(pattern)]
            if i % 2 == 0:
                e = min(L, t + seg)
                s.draw_line((a[0] + ux * t, a[1] + uy * t), (a[0] + ux * e, a[1] + uy * e))
            t += seg
            i += 1
        s.finish(color=(0, 0, 0), width=w, lineCap=0)
        s.commit()

    def poly(self, pts, fill=None, color=None, w=LW, close=True):
        s = self.p.new_shape()
        s.draw_polyline(pts + ([pts[0]] if close else []))
        s.finish(fill=fill, color=color, width=w, closePath=close)
        s.commit()

    def rect(self, r, fill=None, color=None, w=LW):
        s = self.p.new_shape()
        s.draw_rect(fitz.Rect(*r))
        s.finish(fill=fill, color=color, width=w)
        s.commit()

    def oval(self, cx, cy, w, h, fill=(0, 0, 0)):
        s = self.p.new_shape()
        s.draw_oval(fitz.Rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2))
        s.finish(fill=fill, color=None, width=0)
        s.commit()

    def text(self, s, x, y, size, bold=False, align='l', valign='m', color=(0, 0, 0), halo=0.0):
        font = FB if bold else FR
        w = tw(font, s, size)
        x0 = x - w if align == 'r' else x - w / 2 if align == 'c' else x
        cap = size * 0.72                         # 숫자·한글 글자 높이
        base = {'m': y + cap / 2, 'b': y, 't': y + cap}[valign]
        name = 'MB' if bold else 'MR'
        if halo:
            self.p.insert_text((x0, base), s, fontname=name, fontsize=size, render_mode=1,
                               color=(1, 1, 1), border_width=halo / size)
        self.p.insert_text((x0, base), s, fontname=name, fontsize=size, color=color, fill=color)
        return x0, w


def clip_line(s_val, k, xl, xr, yt, yb):
    """x = s - k*(y - yt) 직선을 [xl,xr]x[yt,yb]로 자른 선분."""
    pts = []
    for y in (yt, yb):
        x = s_val - k * (y - yt)
        if xl - 1e-6 <= x <= xr + 1e-6:
            pts.append((x, y))
    for x in (xl, xr):
        y = yt + (s_val - x) / k
        if yt - 1e-6 <= y <= yb + 1e-6:
            pts.append((x, y))
    pts = sorted(set((round(a, 4), round(b, 4)) for a, b in pts))
    return (pts[0], pts[-1]) if len(pts) >= 2 and pts[0] != pts[-1] else None


def symbol(pen, c, yt, yb, ly):
    soil = ly in SOIL
    xl, xr = c - HALF, c + HALF
    if soil:
        for dx in (-SUB, 0, SUB):
            pen.line((c + dx, yt), (c + dx, yb))
        y = yt + 2.14
        while y + 0.8 <= yb:
            pen.oval(c - SUB, y, 1.24, 1.6)
            pen.oval(c + SUB, y, 1.24, 1.6)
            if y + 3.2 + 0.8 <= yb:
                pen.oval(c, y + 3.2, 1.24, 1.6)
            y += 6.4
    else:
        k = 0.777
        kmax = int(math.ceil((xr - xl + k * (yb - yt)) / 5.53)) + 2    # 구간이 길어도 끝까지 빗금
        for kk in range(-3, kmax):
            offs = (0.0, 0.55, 1.1) if ly == '풍화암' else (0.275, 0.825)   # 풍화암 3줄, 연암 2줄 (WR.smf / SR.smf)
            for o in offs:
                seg = clip_line(xl + 2.21 + kk * 5.53 + o, k, xl, xr, yt, yb)
                if seg:
                    pen.line(*seg)
        def plus(x, y):
            if y - 1.065 >= yt and y + 1.065 <= yb:
                pen.line((x, y - 1.065), (x, y + 1.065))
                pen.line((x - 0.83, y), (x + 0.83, y))
        y = yt + 3.56
        while y < yb:
            plus(c - SUB, y)
            plus(c + SUB, y)
            plus(c, y + 3.555)
            y += 7.11


def boundaries(h):
    d = {ly: b for ly, t, b in h['lays']}
    out, cur = [], 0.0
    for ly in ORDER:
        cur = d.get(ly, cur)
        out.append(cur)
    return out


def nfmt(n, p):
    p = int(round(p))
    return f'{n}/ {p}' if p < 10 else f'{n}/{p}'


def spread(ys, gap):
    out = []
    for y in ys:
        out.append(max(y, out[-1] + gap) if out else y)
    return out


def draw(hs, out_pdf, out_jpg=None, dpi=300):
    two = len(hs) == 2
    tops = [h['el'] for h in hs]
    bots = [h['el'] - h['depth'] for h in hs]
    rng = max(tops) - min(bots)
    step = 5.0 if rng > 12 else 2.0
    ax_top = math.ceil(max(tops) / step) * step
    ax_bot = math.floor(min(bots) / step) * step
    if ax_top - max(tops) < 1.0:
        ax_top += step
    v = V if step == 5.0 else min(V * 2.5, 177.8 / (ax_top - ax_bot))
    y_ax_top = 28.3
    Y = lambda el: y_ax_top + (ax_top - el) * v
    y_ax_bot = Y(ax_bot)

    if two:
        dist = math.hypot(hs[1]['x'] - hs[0]['x'], hs[1]['y'] - hs[0]['y'])
        xs = [X1, X1 + max(90.0, min(260.0, dist * H))]
        x_l, x_r = xs
    else:
        x_l, x_r = X1 - 20, X1 + 210
        xs = [(x_l + x_r) / 2 + 30]
    W = x_r + 32 if two else x_r + 12
    Hh = y_ax_bot + (28.8 if two else 12)
    doc = fitz.open()
    page = doc.new_page(width=W, height=Hh)
    pen = Pen(page)

    # ---- 지층 채색
    B = [boundaries(h) for h in hs]
    surf = [Y(h['el']) for h in hs]
    def by(i, k):
        return surf[i] if k < 0 else Y(hs[i]['el'] - B[i][k])
    def edge(k):
        """지층 경계점: 주상 모서리(옆면)에서 그 공의 경계 심도에 맞춘다. 주상 뒤쪽은 수평."""
        if two:
            y0, y1 = by(0, k), by(1, k)
            return [(xs[0], y0), (xs[0] + HALF, y0), (xs[1] - HALF, y1), (xs[1], y1)]
        y = by(0, k)
        return [(x_l, y), (xs[0] - HALF, y), (xs[0] + HALF, y), (x_r, y)]

    def visible(e):          # 주상 바깥에 보이는 선분
        return [(e[1], e[2])] if two else [(e[0], e[1]), (e[2], e[3])]

    last = max(max(ORDER.index(l[0]) for l in h['lays']) for h in hs)
    labels = []
    for k, ly in enumerate(ORDER[:last + 1]):
        top = edge(k - 1)
        bot = edge(k) if k < last else [(p[0], y_ax_bot) for p in top]
        if all(abs(a[1] - b[1]) < 1e-6 for a, b in zip(top, bot)):
            continue
        pen.poly(top + bot[::-1], fill=FILL[ly])
        # 층 이름은 모두 같은 x(두 주상 가운데 / 1공은 왼쪽 가운데)에 세로로 가지런히
        if two:
            cx = (top[1][0] + top[2][0]) / 2
            cy = ((top[1][1] + top[2][1]) / 2 + (bot[1][1] + bot[2][1]) / 2) / 2
            th = (bot[1][1] + bot[2][1]) / 2 - (top[1][1] + top[2][1]) / 2
        else:
            cx = x_l + (xs[0] - HALF - x_l) / 2
            cy, th = (top[1][1] + bot[1][1]) / 2, bot[1][1] - top[1][1]
        labels.append((LABEL.get(ly, ly), cx, cy, th))
    for a, b in visible(edge(-1)):
        pen.line(a, b, w=0.4)
    for k in range(last):
        e, p0 = edge(k), edge(k - 1)
        if all(abs(a[1] - b[1]) < 1e-6 for a, b in zip(e, p0)):
            continue
        for a, b in visible(e):
            pen.dashdot(a, b)
    # 생략 없이 모두 표기: 겹치면 위아래로 최소한만 벌림 (간격 GAP, 순서 유지)
    GAP = 8.6
    labels.sort(key=lambda t: t[2])
    ys = [t[2] for t in labels]
    z = [y - i * GAP for i, y in enumerate(ys)]           # 단조 회귀(PAV)로 최소 이동
    blocks = []
    for zv in z:
        blocks.append([zv, 1])
        while len(blocks) > 1 and blocks[-2][0] / blocks[-2][1] > blocks[-1][0] / blocks[-1][1]:
            v2, n2 = blocks.pop()
            blocks[-1][0] += v2
            blocks[-1][1] += n2
    zz = [b_[0] / b_[1] for b_ in blocks for _ in range(b_[1])]
    for i, (s, cx, cy, th) in enumerate(labels):
        cy = zz[i] + i * GAP
        pen.text(s, cx, cy, 6.71, bold=True, align='c', halo=1.0)

    # ---- 시추주상
    for i, h in enumerate(hs):
        c = xs[i]
        xl, xr = c - HALF, c + HALF
        yt, yb = surf[i], Y(h['el'] - h['depth'])
        pen.rect((xl, yt, xr, yb), fill=(1, 1, 1))
        for ly, t, b in h['lays']:
            symbol(pen, c, Y(h['el'] - t), Y(h['el'] - b), ly)
        pen.line((c, yt), (c, yb))
        for ly, t, b in h['lays'][:-1]:
            pen.line((xl, Y(h['el'] - b)), (xr, Y(h['el'] - b)))
        pen.rect((xl, yt, xr, yb), color=(0, 0, 0), w=0.4)
        # 심도 (왼쪽)
        ybs = [Y(h['el'] - b) for ly, t, b in h['lays']]
        for (ly, t, b), y0, y1 in zip(h['lays'], ybs, spread(ybs, 7.0)):
            pen.line((xl - 3.6, y1), (xl - 1.0, y1) if abs(y1 - y0) > 0.01 else (xl, y0))
            if abs(y1 - y0) > 0.01:
                pen.line((xl - 1.0, y1), (xl, y0))
            pen.text(f'{b:.1f}', xl - 3.8, y1, 5.75, bold=True, align='r', color=DARK, halo=1.4)
        # N값 (오른쪽)
        yns = [Y(h['el'] - dep) for dep, n, p in h['spt']]
        for (dep, n, p), y0, y1 in zip(h['spt'], yns, spread(yns, 6.6)):
            pen.line((xr, y0), (xr + 2.6, y1))
            pen.text(nfmt(n, p), xr + 3.87, y1, 4.8, align='l', halo=0.7)
        # 지반고, 공번
        pen.text(f'EL.(+){h["el"]:.2f}m', c, yt - 2.0, 5.85, align='c', valign='b')
        name = h['name']
        bw = tw(FB, name, 7.67) + 5.0
        yc = yt - 29.27
        pen.line((c, yc + 4.83), (c, yt))
        r = 4.83
        sh = page.new_shape()
        sh.draw_rect(fitz.Rect(c - bw / 2, yc - r, c + bw / 2, yc + r), radius=(r / bw, 0.5))
        sh.finish(fill=(1, 1, 1), color=RED, width=0.8)
        sh.commit()
        pen.text(name, c, yc, 7.67, bold=True, align='c')

    # ---- 표고 축
    pen.text('EL.+m', 21.9, y_ax_top - 15.3, 6.71, bold=True, align='c')
    val = ax_top
    while val >= ax_bot - 1e-6:
        y = Y(val)
        pen.line((11.06, y), (32.32, y))
        pen.text(f'{val:,.1f}', 28.07, y - 0.5, 6.71, bold=True, align='r', valign='b')
        if val - step >= ax_bot - 1e-6:
            pen.rect((30.05, Y(val), 32.32, Y(val - step / 2)), fill=(0, 0, 0))
        val -= step
    pen.line((30.05, y_ax_top), (30.05, y_ax_bot))
    pen.line((32.32, y_ax_top), (32.32, y_ax_bot))

    # ---- 거리
    if two:
        yd = y_ax_bot + 15.6
        a, b = xs
        pen.line((a + 2.4, yd), (b - 2.4, yd), w=0.35)
        pen.rect((a - 0.15, yd - 3.31, a + 0.15, yd + 3.31), fill=(0, 0, 0))
        pen.rect((b - 0.15, yd - 3.31, b + 0.15, yd + 3.31), fill=(0, 0, 0))
        pen.poly([(a + 0.15, yd), (a + 3.22, yd - 1.29), (a + 3.22, yd + 1.29)], fill=(0, 0, 0))
        pen.poly([(b - 0.15, yd), (b - 3.22, yd - 1.29), (b - 3.22, yd + 1.29)], fill=(0, 0, 0))
        pen.text(f'{dist:.2f}m', (a + b) / 2, yd - 1.3, 6.71, bold=True, align='c', valign='b')

    doc.subset_fonts()                      # 쓰는 글자만 글꼴에 포함 (파일 크기 축소)
    doc.save(out_pdf, garbage=4, deflate=True)
    if out_jpg:
        page.get_pixmap(dpi=dpi).save(out_jpg, jpg_quality=95)
    return out_pdf
