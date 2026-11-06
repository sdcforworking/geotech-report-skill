"""4.4.1 시추공별 표준관입시험 표 ↔ 4.3.2 공별 지층표의 N값 범위 대조, 횟수·합계 확인.
사용: python check_spt.py body.txt"""
import re, sys
sys.stdout.reconfigure(encoding='utf-8')
L = open(sys.argv[1], encoding='utf-8').read().split('\n')
body = lambda l: l.split('\t', 1)[1] if '\t' in l else l
strip = lambda s: re.sub(r'⟨\d+×\d+⟩', '', s).strip()

tables, cur = {}, None
for i, l in enumerate(L):
    b = body(l)
    m = re.match(r'┏ 표(\d+) ', b)
    if m:
        cur = int(m.group(1)); tables[cur] = {'line': i + 1, 'rows': []}; continue
    if b.startswith('┗'):
        cur = None; continue
    if cur is not None and b.startswith('┃'):
        tables[cur]['rows'].append([strip(c) for c in re.sub(r'^┃ r\d+: ', '', b).split(' | ')])

def nval(s):
    a, b = map(float, s.split('/'))
    return a if b == 30 else a * 30 / b

# 4.3.2 공별 지층표
s432 = next(i for i, l in enumerate(L) if '(1) 철탑 No.349 구간' in l)
e432 = next(i for i, l in enumerate(L) if body(l).startswith('4.3.3'))
layers = {}
for t, d in tables.items():
    if s432 <= d['line'] <= e432 and d['rows'] and d['rows'][0][0].startswith('시추공번'):
        hole = None
        for r in d['rows'][1:]:
            if r[0]:
                hole = r[0]
            layers.setdefault(hole, []).append(dict(layer=r[1], f=float(r[2]), t=float(r[4]), n=r[7], comp=r[6]))

# 4.4.1 SPT 표
s441 = next(i for i, l in enumerate(L) if body(l).startswith('4.4.1'))
e441 = next(i for i, l in enumerate(L) if body(l).startswith('4.4.2'))
spt, cnt, order = {}, {}, []
for t, d in sorted(tables.items()):
    if not (s441 <= d['line'] <= e441):
        continue
    rows = d['rows']
    if not rows or not rows[0][0].startswith('구 분') or '시험심도' not in ' '.join(rows[0]):
        continue
    hole = None
    for r in rows[2:]:
        if r[0]:
            hole = r[0]; order.append(hole)
            if r[-1]:
                cnt[hole] = int(r[-1])
        off = float(r[1].replace('m', ''))
        for k, v in enumerate(r[2:12]):
            if v not in ('-', ''):
                spt.setdefault(hole, []).append((off + k + 1, v))
print('SPT 표 공 수', len(order), '| 시험 합계', sum(len(v) for v in spt.values()), '| 횟수 칸 합계', sum(cnt.values()))
for h in order:
    if len(spt.get(h, [])) != cnt.get(h):
        print(f'  [횟수] {h}: 값 {len(spt.get(h, []))}개 / 횟수 칸 {cnt.get(h)}')
no_spt = [h for h in layers if h not in spt]
print('SPT 표에 없는 공', len(no_spt), no_spt)
# 공별 지층표 N과 대조
for h, vals in spt.items():
    if h not in layers:
        print(f'  [대조] {h}: 4.3.2에 없음'); continue
    per = {}
    for dep, v in vals:
        lay = next((x for x in layers[h] if x['f'] - 1e-6 <= dep < x['t'] - 1e-6), None)
        if lay is None:
            lay = next((x for x in layers[h] if abs(dep - x['t']) < 1e-6), None)
        if lay is None:
            print(f'  [대조] {h} {dep}m {v}: 해당 지층 없음 (시추심도 넘음?)'); continue
        if lay['layer'] in ('연암', '보통암', '경암'):
            print(f'  [대조] {h} {dep}m {v}: 연암 구간에 SPT 값')
        per.setdefault(id(lay), (lay, []))[1].append(v)
    for lay in layers[h]:
        if lay['layer'] in ('연암', '보통암', '경암'):
            continue
        got = per.get(id(lay), (lay, []))[1]
        if got:
            lo, hi = min(got, key=nval), max(got, key=nval)
            exp = lo if lo == hi else f'{lo}∼{hi}'
        else:
            exp = '-'
        if exp != lay['n']:
            print(f'  [대조] {h} {lay["layer"]}({lay["f"]}∼{lay["t"]}): 4.3.2 "{lay["n"]}" / SPT 표 {got} → "{exp}"')
# 관입불능 정의 확인용: 30cm 미만 50회 개수
allv = [v for vs in spt.values() for _, v in vs]
print('총', len(allv), '| 50/30 미만 관입', sum(1 for v in allv if v.startswith('50/') and float(v.split('/')[1]) < 30))
