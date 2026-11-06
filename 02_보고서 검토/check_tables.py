"""4.3.2 시추조사 결과 자동 대조 — body.txt(hwpx_dump 결과)에서
① 공별 표(심도·두께 연속성) ② 지층두께 총괄표(표68~70) ③ 지층특성 요약표 ④ 지층개황 서술의 숫자가 서로 맞는지 봄.
사용: python check432.py body.txt"""
import re, sys
sys.stdout.reconfigure(encoding='utf-8')
L = open(sys.argv[1], encoding='utf-8').read().split('\n')

def body(l):
    return l.split('\t', 1)[1] if '\t' in l else l

def strip_merge(s):
    return re.sub(r'⟨\d+×\d+⟩', '', s).strip()

# ---- 표 모으기 (표 번호 → 행 목록) + 표가 나온 줄 번호
tables, tline = {}, {}
cur = None
for i, l in enumerate(L):
    b = body(l)
    m = re.match(r'┏ 표(\d+) ', b)
    if m:
        cur = int(m.group(1)); tables[cur] = []; tline[cur] = i + 1
        continue
    if b.startswith('┗'):
        cur = None; continue
    if cur is not None and b.startswith('┃'):
        row = re.sub(r'^┃ r\d+: ', '', b).split(' | ')
        tables[cur].append([strip_merge(c) for c in row])

def num(s):
    s = s.strip()
    return None if s in ('-', '') else float(s)

def nval(s):
    """'50/29' → 51.7, '20/30' → 20"""
    a, b = s.split('/')
    a, b = float(a), float(b)
    return a if b == 30 else a * 30 / b

# ---- 총괄표 (표68~70)
summary = {}
for t in (68, 69, 70):
    for r in tables[t][3:]:
        hole = r[1]
        summary[hole] = dict(모래=num(r[2]), 자갈=num(r[3]), 풍화토=num(r[4]), 풍화암=num(r[5]), 출현=num(r[6]), 심도=num(r[7]))
print('총괄표 공 수', len(summary))
for h, d in summary.items():
    th = sum(v for k, v in d.items() if k in ('모래', '자갈', '풍화토', '풍화암') and v)
    if d['출현'] is not None:
        if abs(th - d['출현']) > 0.05:
            print(f'  [총괄] {h}: 두께 합 {th:.1f} ≠ 기반암 출현 {d["출현"]}')
        if abs(d['심도'] - d['출현'] - 3.0) > 0.05:
            print(f'  [총괄] {h}: 시추심도 {d["심도"]} − 출현 {d["출현"]} = {d["심도"] - d["출현"]:.1f} (연암 3m 아님)')
    else:
        if abs(th - d['심도']) > 0.05:
            print(f'  [총괄] {h}: 두께 합 {th:.1f} ≠ 시추심도 {d["심도"]}')
        if (d['풍화암'] or 0) < 5.0 - 0.05:
            print(f'  [총괄] {h}: 기반암 미출현인데 풍화암 {d["풍화암"]}m (<5m)')

# ---- 구간 나누기 (4.3.2 범위)
start = next(i for i, l in enumerate(L) if body(l).startswith('(1) 철탑 No.349 구간'))
end = next(i for i, l in enumerate(L) if body(l).startswith('4.3.3'))
secs = []
for i in range(start, end):
    m = re.match(r'\((\d+)\) 철탑 (No\.[\d\-]+) 구간', body(L[i]))
    if m:
        secs.append([m.group(2), i])
for k in range(len(secs)):
    secs[k].append(secs[k + 1][1] if k + 1 < len(secs) else end)
print('구간 수', len(secs))

DENS = [(4, '매우느슨'), (10, '느슨'), (30, '보통조밀'), (50, '조밀'), (1e9, '매우조밀')]
def dens(n):
    for lim, w in DENS:
        if n <= lim + 1e-9:
            return w

def fmtrng(a, b, f='{:.1f}'):
    return f.format(a) if abs(a - b) < 1e-9 else f.format(a) + '∼' + f.format(b)

all_holes = set()
for tower, s0, s1 in secs:
    tnums = sorted({int(re.match(r'┏ 표(\d+)', body(L[i])).group(1)) for i in range(s0, s1) if re.match(r'┏ 표(\d+)', body(L[i]))})
    rows, summ = [], None
    for t in tnums:
        rr = tables[t]
        if rr and rr[0][0].startswith('시추공번'):
            hole = None
            for r in rr[1:]:
                if r[0]:
                    hole = r[0]
                rows.append(dict(hole=hole, layer=r[1], f=num(r[2]), t=num(r[4]), th=num(r[5]), comp=r[6], n=r[7]))
        elif rr and rr[0][0].startswith('지 층') and any('출현심도' in c for c in rr[0]):
            summ = (t, rr)
    holes = []
    for r in rows:
        if r['hole'] not in holes:
            holes.append(r['hole'])
    all_holes |= set(holes)
    tag = f'{tower}'
    # A. 공별 표 연속성
    for h in holes:
        hr = [r for r in rows if r['hole'] == h]
        prev = 0.0
        for r in hr:
            if abs(r['f'] - prev) > 0.05:
                print(f'  [{tag} 공별] {h} {r["layer"]}: 시작 {r["f"]} ≠ 앞 지층 끝 {prev}')
            if abs(r['t'] - r['f'] - r['th']) > 0.05:
                print(f'  [{tag} 공별] {h} {r["layer"]}: {r["f"]}∼{r["t"]} 두께 {r["th"]} 계산 {r["t"] - r["f"]:.1f}')
            prev = r['t']
        # B. 총괄표 대조
        S = summary.get(h)
        if S is None:
            print(f'  [{tag}] {h}: 총괄표에 없음'); continue
        cat = {'모래': 0, '자갈': 0, '풍화토': 0, '풍화암': 0}
        rock = None
        for r in hr:
            if r['layer'] == '붕적층':
                cat['자갈' if r['comp'].endswith('자갈') else '모래'] += r['th']
            elif r['layer'] in cat:
                cat[r['layer']] += r['th']
            elif r['layer'] in ('연암', '보통암', '경암'):
                rock = r['f'] if rock is None else rock
            else:
                print(f'  [{tag} 공별] {h}: 지층 이름 "{r["layer"]}"')
        for k, v in cat.items():
            sv = S[k] or 0
            if abs(sv - v) > 0.05:
                print(f'  [{tag}↔총괄] {h} {k}: 공별 {v:.1f} / 총괄 {S[k]}')
        if (rock is None) != (S['출현'] is None) or (rock is not None and abs(rock - S['출현']) > 0.05):
            print(f'  [{tag}↔총괄] {h} 기반암 출현: 공별 {rock} / 총괄 {S["출현"]}')
        if abs(hr[-1]['t'] - S['심도']) > 0.05:
            print(f'  [{tag}↔총괄] {h} 시추심도: 공별 {hr[-1]["t"]} / 총괄 {S["심도"]}')
    # C. 요약표 대조
    if summ is None:
        print(f'  [{tag}] 요약표 못 찾음'); continue
    t, rr = summ
    hdr = rr[0]
    two = hdr[1] == ''          # '지 층' 칸이 두 칸(붕적층 모래/자갈)
    lay_prev = None
    expect = {}
    for r in rr[1:]:
        if two:
            lay = r[0] or lay_prev; sub = r[1]; c = r[2:]
            lay_prev = lay
            key = lay + (f'({sub})' if sub else '')
        else:
            key = r[0]; c = r[1:]
        expect[key] = dict(dep=c[0], th=c[1], comp=c[2], n=c[3])
    for key, e in expect.items():
        lay = key.split('(')[0]; sub = key[len(lay) + 1:-1] if '(' in key else None
        sel = [r for r in rows if r['layer'] == lay and (sub is None or (r['comp'].endswith('자갈') == (sub == '자갈')))]
        if not sel:
            print(f'  [{tag} 요약] {key}: 공별 표에 해당 지층 없음'); continue
        fs = [r['f'] for r in sel]; ths = [r['th'] for r in sel]
        dep_c = fmtrng(min(fs), max(fs)); th_c = fmtrng(min(ths), max(ths))
        if e['dep'] != dep_c:
            print(f'  [{tag} 요약] {key} 출현심도: 표 "{e["dep"]}" / 계산 "{dep_c}"')
        if e['th'] != th_c:
            print(f'  [{tag} 요약] {key} 분포두께: 표 "{e["th"]}" / 계산 "{th_c}"')
        comps = []
        for r in sel:
            for cc in r['comp'].split(', '):
                if cc not in comps:
                    comps.append(cc)
        if lay not in ('연암', '보통암', '경암'):
            if set(e['comp'].split(', ')) != set(comps):
                print(f'  [{tag} 요약] {key} 구성: 표 "{e["comp"]}" / 공별 "{", ".join(comps)}"')
            # N
            ns = []
            for r in sel:
                if r['n'] != '-':
                    ns += [x for x in r['n'].split('∼')]
            if ns:
                lo = min(ns, key=nval); hi = max(ns, key=nval)
                n_c = lo if lo == hi else f'{lo}∼{hi}'
                if e['n'] != n_c:
                    print(f'  [{tag} 요약] {key} N값: 표 "{e["n"]}" / 공별 "{n_c}"')
            elif e['n'] != '-':
                print(f'  [{tag} 요약] {key} N값: 표 "{e["n"]}" / 공별 모두 "-"')
            e['nn'] = ns
        else:
            tc, rq = [], []
            for r in sel:
                m = re.match(r'\[(\d+)/(\d+)\]', r['n'])
                if m:
                    tc.append(int(m.group(1))); rq.append(int(m.group(2)))
            if tc:
                n_c = f'[{fmtrng(min(tc), max(tc), "{}")}/{fmtrng(min(rq), max(rq), "{}")}]'
                if e['n'] != n_c:
                    print(f'  [{tag} 요약] {key} TCR/RQD: 표 "{e["n"]}" / 공별 "{n_c}"')
            e['tc'] = (min(tc), max(tc)) if tc else None; e['rq'] = (min(rq), max(rq)) if rq else None
            rocks = {r['comp'] for r in sel}
            e['rocks'] = rocks
            if len(rocks) > 1:
                print(f'  [{tag} 공별] 연암 암종 섞임: {rocks}')
        e['holes'] = sorted({r['hole'] for r in sel}); e['sel'] = sel
    # D. 지층개황 서술 대조
    paras = {}
    i = s0
    while i < s1:
        b = body(L[i]).strip()
        m = re.match(r'[①-⑨]\s*(.+)$', b)
        if m and i + 1 < s1:
            name = m.group(1).replace(' ', '')
            paras[name] = body(L[i + 1])
            i += 2; continue
        i += 1
    for key, e in expect.items():
        k2 = key.replace(' ', '')
        p = paras.get(k2)
        if p is None:
            print(f'  [{tag} 서술] "{key}" 문단 못 찾음 (있는 제목: {list(paras)})'); continue
        lay = key.split('(')[0]
        # 시추공 언급
        holes_lay = e['holes']
        if '모든 시추공' in p and '제외한' not in p and set(holes_lay) != set(holes):
            print(f'  [{tag} 서술] {key}: "모든 시추공"인데 실제 {holes_lay} / 전체 {holes}')
        mm = re.findall(r'(No\.[\d\-]+\([A-D\-]+\)(?:,\s*No\.[\d\-]+\([A-D\-]+\))*)\s*(?:시추공|를 제외|을 제외)', p)
        if mm and '제외' not in p:
            named = re.findall(r'No\.[\d\-]+\([A-D\-]+\)', mm[0])
            if set(named) != set(holes_lay):
                print(f'  [{tag} 서술] {key}: 서술 공 {named} / 실제 {holes_lay}')
        if '제외' in p:
            ex = re.findall(r'No\.[\d\-]+\([A-D\-]+\)', p.split('제외')[0])
            if set(holes) - set(ex) != set(holes_lay):
                print(f'  [{tag} 서술] {key}: "{",".join(ex)} 제외"인데 실제 {holes_lay}')
        if lay in ('연암', '보통암', '경암'):
            m = re.search(r'지표하\s*([\d.]+(?:∼[\d.]+)?)m', p)
            if not m or m.group(1) != e['dep']:
                print(f'  [{tag} 서술] {key} 출현: 서술 "{m.group(1) if m else None}" / 요약 "{e["dep"]}"')
            m = re.search(r'TCR은\s*([\d∼]+)%,\s*RQD는\s*([\d∼]+)%', p)
            if e.get('tc'):
                tc_c = fmtrng(*e['tc'], '{}'); rq_c = fmtrng(*e['rq'], '{}')
                if not m or m.group(1) != tc_c or m.group(2) != rq_c:
                    print(f'  [{tag} 서술] {key} TCR/RQD: 서술 {m.groups() if m else None} / 공별 ({tc_c}, {rq_c})')
            rk = re.match(r'(\S+?)의 연암', p)
            if rk and e.get('rocks') and {rk.group(1) + '의 연암'} != e['rocks']:
                print(f'  [{tag} 서술] {key}: 서술 "{rk.group(1)}의 연암" / 공별 {e["rocks"]}')
            continue
        m = re.search(r'([\d.]+(?:∼[\d.]+)?)m의\s*층후', p)
        if not m or m.group(1) != e['th']:
            print(f'  [{tag} 서술] {key} 층후: 서술 "{m.group(1) if m else None}" / 요약 "{e["th"]}"')
        m = re.search(r'구성되어', p)
        mc = re.search(r'층후로 분포한다\.\s*(.+?)로 구성되어', p)
        if mc and set(mc.group(1).split(', ')) != set(e['comp'].split(', ')):
            print(f'  [{tag} 서술] {key} 구성: 서술 "{mc.group(1)}" / 요약 "{e["comp"]}"')
        mn = re.search(r'상대밀도는\s*(.+?)로\s*(\S+?)한 상태', p)
        if e.get('nn'):
            if not mn:
                print(f'  [{tag} 서술] {key}: N값 문장 못 찾음'); continue
            txt = mn.group(1)
            got = re.findall(r'(\d+)회/(\d+)cm', txt)
            got_s = '∼'.join(f'{a}/{b}' for a, b in got)
            if got_s != e['n']:
                print(f'  [{tag} 서술] {key} N값: 서술 "{txt}" / 요약 "{e["n"]}"')
            lo = min(e['nn'], key=nval); hi = max(e['nn'], key=nval)
            dl, dh = dens(nval(lo)), dens(nval(hi))
            dw = dl if dl == dh else f'{dl}∼{dh}'
            if mn.group(2) != dw:
                print(f'  [{tag} 서술] {key} 상대밀도: 서술 "{mn.group(2)}" / N {lo}∼{hi} → "{dw}"')
        elif mn and re.search(r'\d+회/', mn.group(1)):
            print(f'  [{tag} 서술] {key}: N값 없음인데 서술에 "{mn.group(1)}"')
        # 위치 표현
        tops = []
        for h in holes_lay:
            hr = [r for r in rows if r['hole'] == h]
            idx = next(j for j, r in enumerate(hr) if r['layer'] == lay and (('(' not in key) or (hr[j]['comp'].endswith('자갈') == key.endswith('(자갈)'))))
            tops.append('지표' if idx == 0 else hr[idx - 1]['layer'])
        pos = re.search(r'(지표면 최상부에서|지표에서부터|지표\S*|\S+ 하부에)', p.split('층후')[0][-40:])
        print(f'  · {tag} {key}: 위 지층 {sorted(set(tops))} / 서술 "{pos.group(1) if pos else "?"}"') if pos and (('하부' in pos.group(1) and '지표' in tops) or ('지표' in pos.group(1) and set(tops) != {'지표'}) or ('하부' in pos.group(1) and pos.group(1).split()[0] not in tops)) else None

print('4.3.2에 나온 공 수', len(all_holes), '| 총괄표에만 있는 공', sorted(set(summary) - all_holes))
