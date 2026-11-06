"""통합 xsm을 새로 읽어 4.2.2 정리 엑셀·공내재하 요약 엑셀의 지층·N값·TCR/RQD·지하수위를 모두 대조 (읽기만 함)."""
import sys, os, re, json, struct
import openpyxl
here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, here)
from xsmparse import parse, strings

XSM = r'C:\Users\심대천\Desktop\동해안-신가평 보고서\1. 최종 정리\01_통합 주상도(xsm)\20261029 동해안~신가평 통합.xsm'
XL = r'C:\Users\심대천\Desktop\동해안-신가평 보고서\1. 최종 정리\05_보고서\4.2.2 시추조사결과_2026.10.27.xlsx'
PMT = r'C:\Users\심대천\Desktop\동해안-신가평 보고서\1. 최종 정리\02_현장시험(공내재하)\공내재하시험 결과 요약(통합).xlsx'

# ── xsm 새로 읽기
X = parse(XSM)
d = open(XSM, 'rb').read(); S = strings(d)
GW = {}
for k, s in enumerate(S):
    if re.fullmatch(r'No\.[\d\-]+\([A-Z\-]+\)', s[2]) and k + 1 < len(S) and S[k + 1][2].endswith('시추'):
        o0 = s[0]
        m1 = next(o for o in range(o0 + 150, o0 + 450) if struct.unpack_from('<d', d, o)[0] == -1.0)
        v = struct.unpack_from('<d', d, m1 - 81)[0]
        GW[s[2]] = None if v == 100000.0 else round(v, 2)
old = json.load(open(os.path.join(here, 'xsm_fin.json'), encoding='utf-8'))
print('xsm 공 수', len(X), '| 이전 추출과 같음', json.loads(json.dumps(X)) == old,
      '| 지하수위 이전 추출과 같음', GW == json.load(open(os.path.join(here, 'gwl.json'), encoding='utf-8')))

tw = lambda h: re.match(r'(No\.[\d\-]+)\(', h).group(1)
f1 = lambda v: f'{v:.1f}'


def lays(h):
    out, top = [], 0.0
    for L in X[h]['lay']:
        out.append(dict(ly=L['ly'], top=round(top, 2), bot=round(L['bot'], 2), comp=L['comp']))
        top = L['bot']
    return out


def spts(h, t, b, last):
    return [(s['n'], s['p']) for s in X[h]['spt'] if s['n'] > 0 and s['p'] > 0 and t - 1e-6 <= s['dep'] and (s['dep'] < b - 1e-6 or (last and s['dep'] <= b + 1e-6))]


def rng(vs, fmt=f1):
    a, b = fmt(min(vs)), fmt(max(vs))
    return a if a == b else f'{a}∼{b}'


def nrange(ss):
    if not ss:
        return '-'
    ss = sorted(ss, key=lambda x: x[0] * 30 / x[1])
    fa, fb = f'{ss[0][0]}/{int(round(ss[0][1]))}', f'{ss[-1][0]}/{int(round(ss[-1][1]))}'
    return fa if fa == fb else f'{fa}∼{fb}'


def crange(h, t, b):
    cs = [(int(c[2][0]), int(c[2][1])) for c in X[h]['core'] if t - 1e-6 <= c[0] < b - 1e-6]
    if not cs:
        return '-'
    return '[' + rng([c[0] for c in cs], str) + '/' + rng([c[1] for c in cs], str) + ']'


norm = lambda s: re.sub(r'\s', '', str(s)).replace('~', '∼') if s is not None else ''
num = lambda v: round(float(v), 2) if isinstance(v, (int, float)) else v
R = {}                 # 시트별 [검사 수, 다른 곳]
def chk(sheet, where, got, exp):
    R.setdefault(sheet, [0, []])
    R[sheet][0] += 1
    if norm(got) != norm(exp) and num(got) != num(exp):
        R[sheet][1].append((where, got, exp))


wb = openpyxl.load_workbook(XL, data_only=True)

# 1) 지층두께 분포현황
ws = wb['지층두께 분포현황']; seen = set()
for r in ws.iter_rows(min_row=7, values_only=True):
    h = r[1]
    if h not in X:
        continue
    seen.add(h); Ls = lays(h)
    th = {k: 0.0 for k in ('모래', '자갈', '풍화토', '풍화암')}
    for L in Ls:
        t = L['bot'] - L['top']
        if L['ly'] == '붕적층':
            th['자갈' if L['comp'].endswith('자갈') else '모래'] += t
        elif L['ly'] in th:
            th[L['ly']] += t
    rock = next((L['top'] for L in Ls if L['ly'] == '연암'), None)
    exp = [f1(th[k]) if th[k] > 0 else '-' for k in ('모래', '자갈', '풍화토', '풍화암')] + [f1(rock) if rock is not None else '-', f1(Ls[-1]['bot'])]
    for nm, g, e in zip(('모래', '자갈', '풍화토', '풍화암', '기반암출현', '시추심도'), r[2:8], exp):
        chk('지층두께 분포현황', f'{h} {nm}', f1(g) if isinstance(g, (int, float)) else g, e)
chk('지층두께 분포현황', '공 수', len(seen), len(X))

# 2) 시추공별 시추조사 결과 (지층·심도·두께·N값/[TCR/RQD])
ws = wb['시추공별 시추조사 결과']; cur = None; got = {}
for r in ws.iter_rows(values_only=True):
    if r[0] in X:
        cur = r[0]; got[cur] = []
    if cur and r[1] and r[1] != '지 층' and isinstance(r[2], (int, float)):
        got[cur].append(r)
    if r[0] and r[0] not in X and r[0] != None and not r[1]:
        cur = None
for h in X:
    Ls = lays(h); g = got.get(h, [])
    chk('시추공별 시추조사 결과', f'{h} 지층 수', len(g), len(Ls))
    for i, (L, row) in enumerate(zip(Ls, g)):
        last = i == len(Ls) - 1
        nv = crange(h, L['top'], L['bot']) if L['ly'] == '연암' else nrange(spts(h, L['top'], L['bot'], last))
        chk('시추공별 시추조사 결과', f'{h} {L["ly"]} 지층', row[1], L['ly'])
        chk('시추공별 시추조사 결과', f'{h} {L["ly"]} 상단', f1(row[2]), f1(L['top']))
        chk('시추공별 시추조사 결과', f'{h} {L["ly"]} 하단', f1(row[4]), f1(L['bot']))
        chk('시추공별 시추조사 결과', f'{h} {L["ly"]} 두께', f1(row[5]), f1(L['bot'] - L['top']))
        chk('시추공별 시추조사 결과', f'{h} {L["ly"]} N값/TCR', row[7], nv)

# 3) 지하수위
ws = wb['지하수위']; seen = set()
for r in ws.iter_rows(min_row=3, values_only=True):
    h = r[1]
    if h not in X:
        continue
    seen.add(h); g = GW[h]
    chk('지하수위', f'{h} 지하수위', r[2], '시추심도 이하' if g is None else f1(g))
    if g is not None:
        ly = next(L['ly'] for L in lays(h) if L['top'] - 1e-6 <= g < L['bot'] - 1e-6)
        chk('지하수위', f'{h} 분포지층', r[3], ly)
chk('지하수위', '공 수', len(seen), len(X))

# 4) 표준관입시험 (공별 1.0~m 칸)
ws = wb['표준관입시험']; cur = None; vals = {}
for r in ws.iter_rows(min_row=3, values_only=True):
    if r[0] in X:
        cur = r[0]; vals[cur] = {}
    if cur and isinstance(r[1], str) and r[1].endswith('m'):
        base = float(r[1][:-1])
        for j, v in enumerate(r[2:12]):
            if v not in (None, '-'):
                vals[cur][round(base + j + 1, 1)] = v
for h in X:
    exp = {round(s['dep'], 1): f"{s['n']}/{int(round(s['p']))}" for s in X[h]['spt'] if s['n'] > 0 and s['p'] > 0}
    chk('표준관입시험', f'{h} 시험 수', len(vals.get(h, {})), len(exp))
    for dep, v in exp.items():
        chk('표준관입시험', f'{h} {dep}m', vals.get(h, {}).get(dep), v)

# 5) 암종별 N값 자료
ws = wb['암종별 N값 자료']; n = 0
spt_all = {(h, round(s['dep'], 1)): f"{s['n']}/{int(round(s['p']))}" for h in X for s in X[h]['spt'] if s['n'] > 0 and s['p'] > 0}
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[0] in ('화강암', '편마암'):
        n += 1
        chk('암종별 N값 자료', f'{r[3]} {r[8]}m', r[9], spt_all.get((r[3], round(r[8], 1))))
chk('암종별 N값 자료', '시험 수', n, len(spt_all))

# 6) 공내재하시험 시트 (4.2.2) + 공내재하 요약 엑셀
def at(h, dep):
    Ls = lays(h)
    L = next((L for L in Ls if L['top'] - 1e-6 <= dep < L['bot'] - 1e-6), Ls[-1])
    if L['ly'] == '연암':
        c = [c for c in X[h]['core'] if c[0] - 1e-6 <= dep <= c[1] + 1e-6][0]
        return L['ly'], f'{c[2][0]}/{c[2][1]}', '-'
    s = [s for s in X[h]['spt'] if abs(s['dep'] - dep) < 1e-6 and s['n'] > 0]
    return L['ly'], '-', f"{s[0]['n']}/{int(round(s[0]['p']))}" if s else '-'


ws = wb['공내재하시험']
for r in ws.iter_rows(min_row=2, values_only=True):
    if r[1] in X:
        ly, tcr, nv = at(r[1], r[2])
        chk('4.2.2 공내재하시험', f'{r[1]} 지층', r[3], ly)
        chk('4.2.2 공내재하시험', f'{r[1]} N값/TCR', r[5], f'[{tcr}]' if ly == '연암' else nv)
pw = openpyxl.load_workbook(PMT, data_only=True)
Sm = pw['요약 상세(참고)']
for r in range(4, 50):
    h = Sm.cell(r, 2).value
    if h not in X:
        continue
    dep = Sm.cell(r, 4).value; ly, tcr, nv = at(h, dep); rs = pw[h]
    chk('공내재하 요약 시트', f'{h} 지층', Sm.cell(r, 5).value, ly)
    chk('공내재하 요약 시트', f'{h} N값/TCR', Sm.cell(r, 6).value, tcr if ly == '연암' else nv)
    chk('공내재하 결과시트', f'{h} 시험심도', rs['D6'].value, dep)
    chk('공내재하 결과시트', f'{h} G.W.L', rs['J4'].value, '시추심도 이하' if GW[h] is None else GW[h])
    chk('공내재하 결과시트', f'{h} Geology', rs['G5'].value, ly)
    chk('공내재하 결과시트', f'{h} TCR/RQD', rs['D7'].value, tcr)
    chk('공내재하 결과시트', f'{h} N-Value', rs['D8'].value, nv)

print()
for k, (n, bad) in R.items():
    print(f'{k}: 검사 {n}칸, 다른 곳 {len(bad)}')
    for b in bad[:30]:
        print('   ', b)
json.dump({k: v for k, v in R.items()}, open(os.path.join(here, 'verify_all.json'), 'w', encoding='utf-8'), ensure_ascii=False, default=str)
