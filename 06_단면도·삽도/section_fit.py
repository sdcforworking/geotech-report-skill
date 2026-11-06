"""보고서 그림칸 크기에 맞춘 지층단면도 (section_ai 양식 그대로, 용지를 칸 크기로 잡고 배치·세로축척만 맞춤).
사용: python section_fit.py 출력폴더 [단면이름 ...]"""
import sys, os, re, json, math
import fitz
import section_ai as sa
from section_ai import Pen, symbol, boundaries, nfmt, spread, ORDER, FILL, LABEL, HALF, RED, DARK, tw, FB

MM = 72 / 25.4
PAD = 1.5            # 칸 안쪽 여유(mm, 사방) — 그림이 칸보다 커서 칸이 늘어나지 않게


def draw_fit(hs, box_w, box_h, out_pdf, out_jpg, dpi=300):
    W, Hp = box_w * MM, box_h * MM
    two = len(hs) == 2
    tops = [h['el'] for h in hs]
    bots = [h['el'] - h['depth'] for h in hs]
    def axis(step):
        t = math.ceil(max(tops) / step) * step
        b = math.floor(min(bots) / step) * step
        if t - max(tops) < 0.5:
            t += step
        return t, b
    for step in (1.0, 2.0, 2.5, 5.0, 10.0):          # 눈금 10칸 이하에서 가장 촘촘한 간격
        ax_top, ax_bot = axis(step)
        if (ax_top - ax_bot) / step <= 10:
            break
    bottom = 19.0 if two else 5.0          # 아래: 거리 치수
    top_min = 21.0                          # 위: 'EL.+m' 글자
    LEAD = 22.0                             # 지표에서 공번 말풍선 중심까지
    head = LEAD + 4.83 + 3.0
    y_ax_top = top_min
    for _ in range(5):
        v = (Hp - y_ax_top - bottom) / (ax_top - ax_bot)
        y_ax_top = max(top_min, head - (ax_top - max(tops)) * v)
    v = (Hp - y_ax_top - bottom) / (ax_top - ax_bot)
    Y = lambda el: y_ax_top + (ax_top - el) * v
    y_ax_bot = Y(ax_bot)

    # 가로: 표고 축(11~32pt) 오른쪽, 왼쪽 공 심도 글자 · 오른쪽 공 N값 자리 확보
    if two:
        dist = math.hypot(hs[1]['x'] - hs[0]['x'], hs[1]['y'] - hs[0]['y'])
        xs = [70.0, W - 36.0]
        x_l, x_r = xs
    else:
        x_l, x_r = 48.0, W - 6.0
        xs = [x_l + (x_r - x_l) * 0.62]
    doc = fitz.open()
    page = doc.new_page(width=W, height=Hp)
    pen = Pen(page)

    B = [boundaries(h) for h in hs]
    surf = [Y(h['el']) for h in hs]

    def by(i, k):
        return surf[i] if k < 0 else Y(hs[i]['el'] - B[i][k])

    def edge(k):
        if two:
            y0, y1 = by(0, k), by(1, k)
            return [(xs[0], y0), (xs[0] + HALF, y0), (xs[1] - HALF, y1), (xs[1], y1)]
        y = by(0, k)
        return [(x_l, y), (xs[0] - HALF, y), (xs[0] + HALF, y), (x_r, y)]

    def visible(e):
        return [(e[1], e[2])] if two else [(e[0], e[1]), (e[2], e[3])]

    last = max(max(ORDER.index(l[0]) for l in h['lays']) for h in hs)
    labels = []
    for k, ly in enumerate(ORDER[:last + 1]):
        top = edge(k - 1)
        bot = edge(k) if k < last else [(p[0], y_ax_bot) for p in top]
        if all(abs(a[1] - b[1]) < 1e-6 for a, b in zip(top, bot)):
            continue
        pen.poly(top + bot[::-1], fill=FILL[ly])
        if two:
            cx = (top[1][0] + top[2][0]) / 2
            cy = ((top[1][1] + top[2][1]) / 2 + (bot[1][1] + bot[2][1]) / 2) / 2
        else:
            cx = x_l + (xs[0] - HALF - x_l) / 2
            cy = (top[1][1] + bot[1][1]) / 2
        labels.append((LABEL.get(ly, ly), cx, cy))
    for a, b in visible(edge(-1)):
        pen.line(a, b, w=0.4)
    for k in range(last):
        e, p0 = edge(k), edge(k - 1)
        if all(abs(a[1] - b[1]) < 1e-6 for a, b in zip(e, p0)):
            continue
        for a, b in visible(e):
            pen.dashdot(a, b)
    GAP = 8.6
    labels.sort(key=lambda t: t[2])
    z = [t[2] - i * GAP for i, t in enumerate(labels)]
    blocks = []
    for zv in z:
        blocks.append([zv, 1])
        while len(blocks) > 1 and blocks[-2][0] / blocks[-2][1] > blocks[-1][0] / blocks[-1][1]:
            v2, n2 = blocks.pop(); blocks[-1][0] += v2; blocks[-1][1] += n2
    zz = [b_[0] / b_[1] for b_ in blocks for _ in range(b_[1])]
    for i, (s, cx, cy) in enumerate(labels):
        cy = min(zz[i] + i * GAP, y_ax_bot - 4.0 - (len(labels) - 1 - i) * GAP)
        pen.text(s, cx, cy, 6.71, bold=True, align='c', halo=1.0)

    for i, h in enumerate(hs):
        c = xs[i]
        xl, xr = c - HALF, c + HALF
        yt, yb = surf[i], Y(h['el'] - h['depth'])
        pen.rect((xl, yt, xr, yb), fill=(1, 1, 1))
        for ly, t, b in h['lays']:
            symbol(pen, c, Y(h['el'] - t), Y(h['el'] - b), ly)
        for ly, t, b in h['lays'][:-1]:
            pen.line((xl, Y(h['el'] - b)), (xr, Y(h['el'] - b)))
        pen.rect((xl, yt, xr, yb), color=(0, 0, 0), w=0.4)
        ybs = [Y(h['el'] - b) for ly, t, b in h['lays']]
        room = y_ax_bot - 1.0
        gd = max(5.2, min(7.0, (room - ybs[0]) / max(1, len(ybs) - 1)))
        for (ly, t, b), y0, y1 in zip(h['lays'], ybs, spread(ybs, gd)):
            pen.line((xl - 3.6, y1), (xl - 1.0, y1) if abs(y1 - y0) > 0.01 else (xl, y0))
            if abs(y1 - y0) > 0.01:
                pen.line((xl - 1.0, y1), (xl, y0))
            pen.text(f'{b:.1f}', xl - 3.8, y1, 5.75, bold=True, align='r', color=DARK, halo=1.4)
        yns = [Y(h['el'] - dep) for dep, n, p in h['spt']]
        if yns:
            gn = min(6.6, (room - yns[0]) / max(1, len(yns) - 1))
            if gn < 4.6:                       # 그래도 모자라면 위로도 벌려서 주상 길이 안에 고르게
                gn = 4.6
                y_start = max(yt - 3.0, room - gn * (len(yns) - 1))
                lab = [min(y0, y_start + i * gn) if False else y_start + i * gn for i, y0 in enumerate(yns)]
            else:
                lab = spread(yns, gn)
            fs = min(4.8, gn * 0.9)
            for (dep, n, p), y0, y1 in zip(h['spt'], yns, lab):
                pen.line((xr, y0), (xr + 2.6, y1))
                pen.text(nfmt(n, p), xr + 3.87, y1, fs, align='l', halo=1.1)
        pen.text(f'EL.(+){h["el"]:.2f}m', c, yt - 2.0, 5.85, align='c', valign='b')
        name = h['name']
        bw = tw(FB, name, 7.67) + 5.0
        yc = yt - LEAD
        pen.line((c, yc + 4.83), (c, yt))
        r = 4.83
        sh = page.new_shape()
        sh.draw_rect(fitz.Rect(c - bw / 2, yc - r, c + bw / 2, yc + r), radius=(r / bw, 0.5))
        sh.finish(fill=(1, 1, 1), color=RED, width=0.8)
        sh.commit()
        pen.text(name, c, yc, 7.67, bold=True, align='c')

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

    if two:
        yd = y_ax_bot + 12.5
        a, b = xs
        pen.line((a + 2.4, yd), (b - 2.4, yd), w=0.35)
        pen.rect((a - 0.15, yd - 3.31, a + 0.15, yd + 3.31), fill=(0, 0, 0))
        pen.rect((b - 0.15, yd - 3.31, b + 0.15, yd + 3.31), fill=(0, 0, 0))
        pen.poly([(a + 0.15, yd), (a + 3.22, yd - 1.29), (a + 3.22, yd + 1.29)], fill=(0, 0, 0))
        pen.poly([(b - 0.15, yd), (b - 3.22, yd - 1.29), (b - 3.22, yd + 1.29)], fill=(0, 0, 0))
        pen.text(f'{dist:.2f}m', (a + b) / 2, yd - 1.3, 6.71, bold=True, align='c', valign='b')

    doc.subset_fonts()
    doc.save(out_pdf, garbage=4, deflate=True)
    pix = page.get_pixmap(dpi=dpi)
    pix.set_dpi(dpi, dpi)                     # 한글에 넣으면 칸 크기(mm) 그대로 들어가도록
    pix.save(out_jpg, jpg_quality=95)


if __name__ == '__main__':
    import section_dxf as sd
    here = os.path.dirname(os.path.abspath(__file__))
    out = sys.argv[1]
    only = sys.argv[2:]
    H = sd.load(os.path.join(here, 'xsm5.json'), 'C:/Users/심대천/Desktop/동해안-신가평 보고서/좌표.xlsx')
    secs = dict(json.load(open(os.path.join(here, 'sections_cur.json'), encoding='utf-8')))
    cm = json.load(open(os.path.join(here, 'cellmap.json'), encoding='utf-8'))
    os.makedirs(os.path.join(out, 'ai'), exist_ok=True); os.makedirs(os.path.join(out, 'jpg'), exist_ok=True)
    rows = []
    for twr, c in cm.items():
        for nm, n in c['secs']:
            if only and nm not in only:
                continue
            bw, bh = round(c['w'] - 2 * PAD, 1), round(c['h'] - 2 * PAD, 1)
            draw_fit([H[h] for h in secs[nm]], bw, bh,
                     os.path.join(out, 'ai', f'지층단면도_{nm}.ai'), os.path.join(out, 'jpg', f'지층단면도_{nm}.jpg'))
            rows.append((nm, c['w'], c['h'], bw, bh))
    print(len(rows), 'done')
    json.dump(rows, open(os.path.join(out, 'sizes.json'), 'w', encoding='utf-8'), ensure_ascii=False)
