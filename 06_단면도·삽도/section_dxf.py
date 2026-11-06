"""지층단면도 DXF 생성 (참고: 종단면도 삽도 참고자료 / 지층단면도_No.1(273)-A.jpg 양식).

단위: 도면 mm (삽도용, 실축척 아님)
자료: xsm에서 뽑은 지층·SPT(json), 좌표.xlsx(좌표·지반고)
"""
import json, math, sys, collections
import ezdxf
import openpyxl
from ezdxf.enums import TextEntityAlignment as TA

# ---- 양식 값
FILL = {'붕적층': (237, 227, 214), '풍화토': (188, 181, 167), '풍화암': (169, 163, 151), '연암': (158, 145, 138)}
ORDER = ['붕적층', '풍화토', '풍화암', '연암']
LABEL = {'연암': '연암'}
RED = (230, 0, 18)
TH_N, TH_D, TH_EL, TH_AX, TH_NAME, TH_LAY, TH_DIST = 2.0, 2.2, 2.3, 2.6, 3.4, 3.4, 3.0
COL_W = 4.5          # 주상 폭
X0 = 38.0            # 첫 시추공 x (축 오른쪽)


def load(xsm_json, coord_xlsx):
    H = json.load(open(xsm_json, encoding='utf-8'))
    ws = openpyxl.load_workbook(coord_xlsx, data_only=True)['시추조사 위치현황']
    C = {r[0]: (float(r[1]), float(r[2]), float(r[3])) for r in ws.iter_rows(min_row=2, values_only=True) if r[0]}
    holes = {}
    for h, v in H.items():
        top, lays = 0.0, []
        for L in v['lay']:
            lays.append((L['ly'], round(top, 2), round(L['bot'], 2)))
            top = L['bot']
        spt = [(s['dep'], s['n'], s['p']) for s in v['spt'] if s['n'] > 0 and s['p'] > 0]
        x, y, el = C[h]
        holes[h] = dict(name=h, lays=lays, spt=spt, x=x, y=y, el=el, depth=lays[-1][2])
    return holes


def nfmt(n, p):
    p = int(round(p))
    return f'{n}/{p:>2d}'.replace(' ', ' ') if p < 10 else f'{n}/{p}'


def spread(ys, gap):
    """위에서부터 글자 y를 최소 간격 gap 이상으로 벌린다 (순서 유지)."""
    out = []
    for y in ys:
        out.append(min(y, out[-1] - gap) if out else y)
    return out


def boundaries(h):
    """ORDER 순서의 각 지층 하단 심도 (없는 지층은 두께 0으로 처리)."""
    d = {ly: (t, b) for ly, t, b in h['lays']}
    out, cur = [], 0.0
    for ly in ORDER:
        if ly in d:
            cur = d[ly][1]
        out.append(cur)
    return out          # 붕적층 하단, 풍화토 하단, 풍화암 하단, 연암 하단


def setup(doc):
    doc.styles.new('KOR', dxfattribs={'font': 'malgun.ttf'})
    doc.styles.new('KORB', dxfattribs={'font': 'malgunbd.ttf'})
    for name, col in [('지층채색', 8), ('지층경계', 7), ('주상도', 7), ('글자', 7), ('표고축', 7), ('공번', 1), ('치수', 7)]:
        doc.layers.add(name, color=col)
    if 'DASHDOT' not in doc.linetypes:
        doc.linetypes.add('DASHDOT', pattern=[6.0, 4.0, -1.2, 0.0, -1.2], description='Dash dot')


def text(msp, s, x, y, h, align, layer='글자', style='KOR', color=None):
    at = {'layer': layer, 'height': h, 'style': style}
    if color:
        at['true_color'] = ezdxf.colors.rgb2int(color)
    msp.add_text(s, dxfattribs=at).set_placement((x, y), align=align)


def draw(holes_in, out, title=None):
    hs = holes_in
    doc = ezdxf.new('R2010', setup=True)
    setup(doc)
    msp = doc.modelspace()

    # ---- 축척
    tops = [h['el'] for h in hs]
    bots = [h['el'] - h['depth'] for h in hs]
    rng = max(tops) - min(bots)
    step = 5.0 if rng > 14 else 2.0
    ax_top = math.ceil((max(tops) + 0.5) / step) * step
    ax_bot = math.floor((min(bots) - 1.5) / step) * step
    vs = max(3.4, min(6.0, 120.0 / (ax_top - ax_bot)))          # mm per m
    Y = lambda el: (el - ax_bot) * vs
    if len(hs) == 1:
        width = 120.0
        xs = [X0 + width / 2]
        x_l, x_r = X0, X0 + width
    else:
        dist = math.hypot(hs[1]['x'] - hs[0]['x'], hs[1]['y'] - hs[0]['y'])
        width = max(70.0, min(140.0, 40 + 3.2 * dist))
        xs = [X0, X0 + width]
        x_l, x_r = xs
    y_bot = Y(ax_bot)

    # ---- 지층 채색과 경계선
    B = [boundaries(h) for h in hs]
    surf = [Y(h['el']) for h in hs]

    def bound_y(i, k):          # k = -1 지표, 0..3 지층 하단
        return surf[i] if k < 0 else Y(hs[i]['el'] - B[i][k])

    def pts(k):
        if len(hs) == 1:
            return [(x_l, bound_y(0, k)), (x_r, bound_y(0, k))]
        return [(xs[0], bound_y(0, k)), (xs[1], bound_y(1, k))]

    last = max(max(i for i, ly in enumerate(ORDER) if any(l[0] == ly for l in h['lays'])) for h in hs)
    for k, ly in enumerate(ORDER):
        if k > last:
            break
        top = pts(k - 1)
        bot = pts(k) if k < last else [(top[0][0], y_bot), (top[1][0], y_bot)]
        if all(abs(a[1] - b[1]) < 1e-6 for a, b in zip(top, bot)):
            continue
        poly = [top[0], top[1], bot[1], bot[0]]
        hatch = msp.add_hatch(dxfattribs={'layer': '지층채색'})
        hatch.rgb = FILL[ly]
        hatch.paths.add_polyline_path(poly, is_closed=True)
        # 층 이름: 가운데
        cx = (x_l + x_r) / 2 if len(hs) == 2 else x_l + (x_r - x_l) * 0.22
        ty = (top[0][1] + top[1][1]) / 2
        by = (bot[0][1] + bot[1][1]) / 2
        if ty - by > TH_LAY * 1.2:
            text(msp, LABEL.get(ly, ly), cx, (ty + by) / 2, TH_LAY, TA.MIDDLE_CENTER, style='KORB')
    # 지표선(실선)과 지층 경계(1점쇄선)
    msp.add_line(*pts(-1), dxfattribs={'layer': '지층경계'})
    for k in range(min(last, 3)):
        a, b = pts(k)
        if abs(a[1] - pts(k - 1)[0][1]) < 1e-6 and abs(b[1] - pts(k - 1)[1][1]) < 1e-6:
            continue
        msp.add_line(a, b, dxfattribs={'layer': '지층경계', 'linetype': 'DASHDOT', 'ltscale': 0.6})

    # ---- 시추주상
    for i, h in enumerate(hs):
        x = xs[i]
        xl, xr = x - COL_W / 2, x + COL_W / 2
        ytop, ybot = surf[i], Y(h['el'] - h['depth'])
        bg = msp.add_hatch(dxfattribs={'layer': '주상도'})
        bg.rgb = (255, 255, 255)
        bg.paths.add_polyline_path([(xl, ytop), (xr, ytop), (xr, ybot), (xl, ybot)], is_closed=True)
        for ly, t, b in h['lays']:
            y1, y2 = Y(h['el'] - t), Y(h['el'] - b)
            if ly in ('붕적층', '풍화토'):
                # 점 3열
                n = int((y1 - y2) / 1.6)
                for j in range(n):
                    yy = y1 - 0.8 - j * 1.6
                    for dx in (-1.4, 0.0, 1.4):
                        off = 0.4 if (j % 2 and dx == 0.0) else 0.0
                        ht = msp.add_hatch(dxfattribs={'layer': '주상도'})
                        ht.rgb = (0, 0, 0)
                        ht.paths.add_edge_path().add_arc((x + dx, yy - off), 0.28, 0, 360)
            else:
                hat = msp.add_hatch(dxfattribs={'layer': '주상도'})
                hat.set_pattern_fill('ANSI31', scale=0.35, color=7)
                hat.paths.add_polyline_path([(xl, y1), (xr, y1), (xr, y2), (xl, y2)], is_closed=True)
                yy = y1 - 1.2
                while yy > y2 + 0.6:
                    for dx in (-1.1, 1.1):
                        msp.add_line((x + dx - 0.5, yy), (x + dx + 0.5, yy), dxfattribs={'layer': '주상도'})
                        msp.add_line((x + dx, yy - 0.5), (x + dx, yy + 0.5), dxfattribs={'layer': '주상도'})
                    yy -= 2.2
        msp.add_lwpolyline([(xl, ytop), (xr, ytop), (xr, ybot), (xl, ybot)], close=True,
                           dxfattribs={'layer': '주상도', 'const_width': 0.25})
        # 지층 경계 심도(왼쪽)
        ybs = [Y(h['el'] - b) for ly, t, b in h['lays']]
        for (ly, t, b), yb, yt in zip(h['lays'], ybs, spread(ybs, TH_D * 1.35)):
            msp.add_line((xl - 2.0, yt), (xl, yb), dxfattribs={'layer': '주상도'}) if abs(yt - yb) > 1e-6 else None
            msp.add_line((xl, yb), (xr, yb), dxfattribs={'layer': '주상도'})
            if abs(yt - yb) < 1e-6:
                msp.add_line((xl - 2.0, yb), (xl, yb), dxfattribs={'layer': '주상도'})
            text(msp, f'{b:.1f}', xl - 2.4, yt, TH_D, TA.MIDDLE_RIGHT, style='KORB')
        # N값(오른쪽)
        yns = [Y(h['el'] - dep) for dep, n, p in h['spt']]
        for (dep, n, p), yy, yt in zip(h['spt'], yns, spread(yns, TH_N * 1.15)):
            msp.add_line((xr, yy), (xr + 1.2, yt), dxfattribs={'layer': '주상도'})
            text(msp, nfmt(n, p), xr + 1.6, yt, TH_N, TA.MIDDLE_LEFT)
        # 공번 상자, 지반고
        yn = max(surf) + 14.0 + (0 if i == 0 or len(hs) == 1 else 0)
        yname = ytop + 12.0
        name = h['name']
        wbox = len(name) * TH_NAME * 0.62 + 4.0
        r = 2.6
        box = msp.add_lwpolyline([(x - wbox / 2, yname - r, 0), (x + wbox / 2, yname - r, 1),
                                  (x + wbox / 2, yname + r, 0), (x - wbox / 2, yname + r, 1)], format='xyb',
                                 close=True, dxfattribs={'layer': '공번', 'const_width': 0.3})
        box.dxf.true_color = ezdxf.colors.rgb2int(RED)
        text(msp, name, x, yname, TH_NAME, TA.MIDDLE_CENTER, style='KORB')
        msp.add_line((x, yname - 2.6), (x, ytop), dxfattribs={'layer': '공번'})
        text(msp, f'EL.(+){h["el"]:.2f}m', x, ytop + 4.0, TH_EL, TA.BOTTOM_CENTER)

    # ---- 표고 축
    ax = 6.0
    text(msp, 'EL.+m', ax, Y(ax_top) + 8.0, TH_AX, TA.BOTTOM_CENTER, style='KORB')
    v = ax_top
    i = 0
    while v >= ax_bot - 1e-6:
        yy = Y(v)
        msp.add_line((ax - 9, yy), (ax + 2.5, yy), dxfattribs={'layer': '표고축'})
        text(msp, f'{v:.1f}', ax - 1, yy + 0.6, TH_AX, TA.BOTTOM_RIGHT, style='KORB')
        if v - step >= ax_bot - 1e-6:
            y2 = Y(v - step)
            mid = (yy + y2) / 2
            if i % 1 == 0:
                bar = msp.add_hatch(dxfattribs={'layer': '표고축'})
                bar.rgb = (0, 0, 0)
                bar.paths.add_polyline_path([(ax + 0.5, yy), (ax + 2.5, yy), (ax + 2.5, mid), (ax + 0.5, mid)], is_closed=True)
        v -= step
        i += 1
    msp.add_lwpolyline([(ax + 0.5, Y(ax_top)), (ax + 2.5, Y(ax_top)), (ax + 2.5, Y(ax_bot)), (ax + 0.5, Y(ax_bot))],
                       close=True, dxfattribs={'layer': '표고축'})

    # ---- 거리 치수
    if len(hs) == 2:
        yd = y_bot - 7.0
        msp.add_line((xs[0], yd), (xs[1], yd), dxfattribs={'layer': '치수'})
        for x, sgn in ((xs[0], 1), (xs[1], -1)):
            msp.add_line((x, yd - 2), (x, yd + 2), dxfattribs={'layer': '치수'})
            s = msp.add_solid([(x, yd), (x + sgn * 2.2, yd + 0.6), (x + sgn * 2.2, yd - 0.6)], dxfattribs={'layer': '치수'})
        text(msp, f'{dist:.2f}m', (xs[0] + xs[1]) / 2, yd + 1.2, TH_DIST, TA.BOTTOM_CENTER, style='KORB')
    if title:
        doc.header['$PROJECTNAME'] = title
    doc.saveas(out)
    return doc
