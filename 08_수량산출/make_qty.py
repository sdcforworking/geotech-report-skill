"""이번 과업(동해안~신가평 10공구) 수량산출서 만들기 — 예시 '수량산출서.xlsx'의 시추·시험 수량산출(수행) 양식을 따름.
자료: 본문 hwpx(2026.11.02(2))를 풀어 쓴 body.txt(지층두께 총괄표·SPT·공내재하·토질시험 표), 공별 규격은 통합 주상도 xsm에서 뽑은 size.json.
병합 칸 없이 만들고(제목은 '선택 영역 가운데'), 합계는 함수로 둠.
사용: python make_qty.py body.txt size.json 출력.xlsx"""
import re, sys, json
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter as CL

sys.stdout.reconfigure(encoding='utf-8')
L = open(sys.argv[1], encoding='utf-8').read().split('\n')
SIZE = json.load(open(sys.argv[2], encoding='utf-8'))
OUT = sys.argv[3]
body = lambda l: l.split('\t', 1)[1] if '\t' in l else l
strip = lambda s: re.sub(r'⟨\d+×\d+⟩', '', s).strip()

tables, cur = {}, None
for i, l in enumerate(L):
    b = body(l)
    m = re.match(r'┏ 표(\d+) ', b)
    if m:
        cur = int(m.group(1)); tables[cur] = {'line': i, 'rows': []}; continue
    if b.startswith('┗'):
        cur = None; continue
    if cur is not None and b.startswith('┃'):
        tables[cur]['rows'].append([strip(c) for c in re.sub(r'^┃ r\d+: ', '', b).split(' | ')])
sec = lambda name: next(i for i, l in enumerate(L) if body(l).startswith(name))
num = lambda s: 0.0 if s in ('-', '') else float(s)

# 지층두께 총괄표 (4.3.2 첫 표들) — 철탑, 공번, 모래, 자갈, 풍화토, 풍화암, 기반암 출현, 시추심도
holes, tower = [], None
s432, s433 = sec('4.3.2'), sec('4.3.3')
for t, d in sorted(tables.items()):
    if s432 < d['line'] < s433 and d['rows'] and d['rows'][0][0].startswith('철탑') and '기반암' in ' '.join(d['rows'][0]):
        for r in d['rows'][3:]:
            tower = r[0] or tower
            sand, grav, ws, wr, top, dep = map(num, r[2:8])
            soft = round(dep - top, 2) if r[6] not in ('-', '') else 0.0
            assert abs(sand + grav + ws + wr + soft - dep) < 0.05, r
            holes.append(dict(tower=tower, hole=r[1], sand=sand, gravel=grav, ws=ws, wr=wr, soft=soft, depth=dep))
assert len(holes) == 122, len(holes)

# 표준관입시험 횟수 (4.4.1), 공내재하 (4.4.2), 토질시험 (4.5.1)
spt, pmt, lab = {}, {}, {}
s441, s442, s451, s5 = sec('4.4.1'), sec('4.4.2'), sec('4.5.1'), sec('제5장')
for t, d in sorted(tables.items()):
    rows = d['rows']
    if not rows:
        continue
    if s441 < d['line'] < s442 and rows[0][0].startswith('구 분') and '시험심도' in ' '.join(rows[0]):
        for r in rows[2:]:
            if r[0] and r[-1]:
                spt[r[0]] = int(r[-1])
    if s442 < d['line'] < s451 and rows[0][0] == '철탑번호':
        for r in rows[1:]:
            pmt[r[1]] = pmt.get(r[1], 0) + 1
    if s451 < d['line'] < s5 and rows[0][0].startswith('시추'):
        for r in rows[2:]:
            if r[0].startswith('No.'):
                lab[r[0]] = lab.get(r[0], 0) + 1
print('SPT', sum(spt.values()), len(spt), '| 공내재하', sum(pmt.values()), '| 토질', sum(lab.values()))
assert sum(spt.values()) == 720 and sum(pmt.values()) == 46 and sum(lab.values()) == 46
assert set(SIZE) == {h['hole'] for h in holes}

# ---------- 서식 ----------
F = Font(name='맑은 고딕', size=10); FB = Font(name='맑은 고딕', size=10, bold=True)
FT = Font(name='맑은 고딕', size=14, bold=True)
thin = Side(style='thin', color='808080'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD = PatternFill('solid', fgColor='D9E1F2'); SUMF = PatternFill('solid', fgColor='FFF2CC')
C = Alignment(horizontal='center', vertical='center', wrap_text=True)
CC = Alignment(horizontal='centerContinuous', vertical='center', wrap_text=True)
NUM1 = '#,##0.0_);[Red]\\(#,##0.0\\)'; NUM0 = '#,##0_);[Red]\\(#,##0\\)'


def cell(ws, r, c, v, font=F, fill=None, fmt=None, align=C, border=True):
    x = ws.cell(r, c, v); x.font = font; x.alignment = align
    if border: x.border = BOX
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    return x


def span_title(ws, r, c1, c2, text, font=FB, fill=HEAD):
    for c in range(c1, c2 + 1):
        cell(ws, r, c, text if c == c1 else None, font, fill, align=CC)


wb = Workbook()

# ---------- 시추 수량산출(수행) ----------
ws = wb.active; ws.title = '시추 수량산출(수행)'
# 열: A 공번, B 규격, C 모래 D 점토 E 자갈 F 풍화토 G 풍화암 H 연암 I 경암, J 계, K 보링공보호, L 비고(철탑), M 구분(육상)
for c in range(1, 14):
    ws.cell(1, c).alignment = CC
ws.cell(1, 1, '시 추 수 량 산 출 서 (수행)').font = FT
hdr2 = ['구 분', '규격', '지층별 심도(m)', None, None, None, None, None, None, '계', '보링공보호\n(연,경암층 제외)', '비 고', '구분']
for c, h in enumerate(hdr2, 1):
    cell(ws, 2, c, h, FB, HEAD)
for c in range(3, 10):
    ws.cell(2, c).alignment = CC
for c, h in enumerate(['모래층', '점토층', '자갈층', '풍화토층', '풍화암층', '연암층', '경암층'], 3):
    cell(ws, 3, c, h, FB, HEAD)
for c in (1, 2, 10, 11, 12, 13):
    cell(ws, 3, c, None, FB, HEAD)
r0 = 4
for i, h in enumerate(holes):
    r = r0 + i
    vals = [h['hole'], SIZE[h['hole']], h['sand'], 0, h['gravel'], h['ws'], h['wr'], h['soft'], 0]
    for c, v in enumerate(vals, 1):
        cell(ws, r, c, v, fmt=NUM1 if c >= 3 else None)
    cell(ws, r, 10, f'=SUM(C{r}:I{r})', fmt=NUM1)
    cell(ws, r, 11, f'=J{r}-I{r}-H{r}', fmt=NUM1)
    cell(ws, r, 12, h['tower']); cell(ws, r, 13, '육상')
rl = r0 + len(holes) - 1
rs = rl + 1
for k, (lab_, sz) in enumerate([('BX 합계', 'BX'), ('NX 합계', 'NX')]):
    r = rs + k
    cell(ws, r, 1, '육상' if k == 0 else None, FB, SUMF); cell(ws, r, 2, lab_, FB, SUMF)
    for c in range(3, 10):
        L_ = CL(c)
        cell(ws, r, c, f'=SUMPRODUCT(($B${r0}:$B${rl}="{sz}")*($M${r0}:$M${rl}="육상"),{L_}{r0}:{L_}{rl})', FB, SUMF, NUM1)
    cell(ws, r, 10, f'=SUM(C{r}:I{r})', FB, SUMF, NUM1)
    cell(ws, r, 11, f'=SUM(C{r}:G{r})', FB, SUMF, NUM1)
    cell(ws, r, 12, f'=COUNTIFS($B${r0}:$B${rl},"{sz}",$M${r0}:$M${rl},"육상")', FB, SUMF, '0"공"')
    cell(ws, r, 13, None, FB, SUMF)
rt = rs + 2
cell(ws, rt, 1, '합계', FB, SUMF); cell(ws, rt, 2, None, FB, SUMF)
for c in range(3, 12):
    L_ = CL(c); cell(ws, rt, c, f'=SUM({L_}{rs}:{L_}{rs + 1})', FB, SUMF, NUM1)
cell(ws, rt, 12, f'=L{rs}+L{rs + 1}', FB, SUMF, '0"공"'); cell(ws, rt, 13, None, FB, SUMF)
note = rt + 1
ws.cell(note, 1, '※ 모래층·자갈층 = 붕적층(모래·자갈), 풍화토층 = 풍화토. 점토층·경암층은 이번 과업에서 나오지 않음. '
                 '시추 규격(BX/NX)은 통합 주상도(xsm), 지층 두께는 본문 4.3.2 지층두께 분포현황 기준.').font = F
rp = note + 2
ws.cell(rp, 1, '폐공수량 산출(육상)').font = FB
items = [('  - 시추공 폐공 (육상) / (토사,풍화암) BX : ', f'=K{rs}', 'm'),
         ('  - 시추공 폐공 (육상) / (연암,경암) BX : ', f'=H{rs}+I{rs}', 'm'),
         ('  - 시추공 폐공 (육상) / (토사,풍화암) NX : ', f'=K{rs + 1}', 'm'),
         ('  - 시추공 폐공 (육상) / (연암,경암) NX : ', f'=H{rs + 1}+I{rs + 1}', 'm'),
         ('  -  Boring공 보호(연암, 경암층 제외) :', f'=K{rt}', 'm'),
         ('  -  표준관입시험 (연암 상부 토사·풍화암, 1.0m 간격) :', None, '회')]
for k, (t, f, u) in enumerate(items):
    r = rp + 1 + k
    ws.cell(r, 1, t).font = F
    x = ws.cell(r, 7, f); x.font = FB; x.number_format = NUM1 if u == 'm' else NUM0
    ws.cell(r, 8, u).font = F
spt_ref_row = rp + 6
for c, w in zip('ABCDEFGHIJKLM', [14, 7, 9, 9, 9, 10, 10, 9, 9, 10, 15, 11, 7]):
    ws.column_dimensions[c].width = w
ws.row_dimensions[2].height = 30
ws.freeze_panes = 'C4'
qty_sheet = ws

# ---------- 시험 수량산출(수행) ----------
ws = wb.create_sheet('시험 수량산출(수행)')
for c in range(1, 11):
    ws.cell(1, c).alignment = CC
ws.cell(1, 1, '실내 · 실외시험 수행내역서').font = FT
groups = [('구 분', 1, 1), ('실내시험(토사층)', 2, 6), ('실외시험', 7, 9), ('비 고', 10, 10)]
for t, a, b in groups:
    span_title(ws, 2, a, b, t)
for c, h in enumerate(['', '함수량', '비중', '액성한계', '소성한계', '입도분포', '표준관입', '공내재하', '지하수위', ''], 1):
    cell(ws, 3, c, h or None, FB, HEAD)
t0 = 4
for i, h in enumerate(holes):
    r = t0 + i; n = lab.get(h['hole'], 0)
    vals = [h['hole'], n, n, n, n, n, spt.get(h['hole'], 0), pmt.get(h['hole'], 0), 1, h['tower']]
    for c, v in enumerate(vals, 1):
        cell(ws, r, c, v, fmt=NUM0 if 2 <= c <= 9 else None)
tl = t0 + len(holes) - 1
tr = tl + 1
cell(ws, tr, 1, '합계', FB, SUMF)
for c in range(2, 10):
    L_ = CL(c); cell(ws, tr, c, f'=SUM({L_}{t0}:{L_}{tl})', FB, SUMF, NUM0)
cell(ws, tr, 10, None, FB, SUMF)
ws.cell(tr + 1, 1, '※ 토질시험은 시료 1개당 함수량·비중·액성한계·소성한계·입도분포 각 1회. 지하수위는 시추 완료 72시간 뒤 공별 1회 측정.').font = F
ws.cell(tr + 2, 1, '※ 공별로 나누지 않는 항목: 굴절법 탄성파탐사 2측선(S-1 46m, S-2 46m) 총 0.092km.').font = F
for c, w in zip('ABCDEFGHIJ', [14, 8, 8, 9, 9, 9, 9, 9, 9, 10]):
    ws.column_dimensions[c].width = w
ws.freeze_panes = 'B4'
qty_sheet.cell(spt_ref_row, 7).value = f"='시험 수량산출(수행)'!G{tr}"

# ---------- 수량 총괄 (맨 앞) ----------
ws = wb.create_sheet('수량 총괄', 0)
for c in range(1, 9):
    ws.cell(1, c).alignment = CC
ws.cell(1, 1, '수 량 총 괄 표').font = FT
for c, h in enumerate(['구 분', '세 부', '규 격', '단 위', '도 급', '수 행', '증 감', '비 고'], 1):
    cell(ws, 2, c, h, FB, HEAD)
S = "'시추 수량산출(수행)'"; T = "'시험 수량산출(수행)'"
rows = [
    ('시추조사', 'BX규격', '-', '공', 70, f'={S}!L{rs}', ''),
    (None, 'NX규격', '-', '공', 48, f'={S}!L{rs + 1}', ''),
    (None, '합 계', '-', '공', 118, f'={S}!L{rt}', ''),
    (None, '시추 연장 (BX)', '토사·풍화암 / 연암', 'm', '-', f'={S}!J{rs}', '토사·풍화암 {} m, 연암 {} m'),
    (None, '시추 연장 (NX)', '토사·풍화암 / 연암', 'm', '-', f'={S}!J{rs + 1}', '토사·풍화암 {} m, 연암 {} m'),
    (None, 'Boring공 보호', '연암 제외', 'm', '-', f'={S}!K{rt}', ''),
    ('현장시험', '표준관입시험', 'KS F 2307', '회', 649, f'={T}!G{tr}', ''),
    (None, '공내재하시험', '-', '회', 48, f'={T}!H{tr}', ''),
    (None, '지하수위 측정', '-', '공', '-', f'={T}!I{tr}', ''),
    ('물리탐사', '굴절법탄성파탐사', '-', 'km', '-', 0.092, '2측선 (S-1 46m, S-2 46m)'),
    ('토질시험', '비중', 'KS F 2308', '회', 48, f'={T}!C{tr}', ''),
    (None, '함수비', 'KS F 2306', '회', 48, f'={T}!B{tr}', ''),
    (None, '액성한계', 'KS F 2303', '회', 48, f'={T}!D{tr}', ''),
    (None, '소성한계', 'KS F 2303', '회', 48, f'={T}!E{tr}', ''),
    (None, '입도분석', 'KS F 2302', '회', 48, f'={T}!F{tr}', ''),
]
for k, (a, b, c3, d, e, f, g) in enumerate(rows):
    r = 3 + k
    unit_m = d in ('m', 'km')
    fmt = '#,##0.000' if d == 'km' else (NUM1 if unit_m else NUM0)
    cell(ws, r, 1, a, FB); cell(ws, r, 2, b); cell(ws, r, 3, c3); cell(ws, r, 4, d)
    cell(ws, r, 5, e, fmt=fmt); cell(ws, r, 6, f, fmt=fmt)
    cell(ws, r, 7, f'=IF(ISNUMBER(E{r}),F{r}-E{r},"-")', fmt=fmt)
    if '{}' in g:
        rr = rs if 'BX' in b else rs + 1
        g = f'="토사·풍화암 "&TEXT({S}!K{rr},"#,##0.0")&" m, 연암 "&TEXT({S}!H{rr},"#,##0.0")&" m"'
    cell(ws, r, 8, g or None, align=Alignment(horizontal='left', vertical='center'))
ws.cell(3 + len(rows), 1, '※ 도급 수량은 본문 1.3 조사항목 및 수량 표 기준. 시추 연장·지하수위는 도급 수량이 없어 "-"로 둠.').font = F
for c, w in zip('ABCDEFGH', [11, 18, 18, 6, 9, 11, 9, 34]):
    ws.column_dimensions[c].width = w

for s in wb.worksheets:
    s.sheet_view.tabSelected = False
wb.active = 0; wb.worksheets[0].sheet_view.tabSelected = True
wb.save(OUT)
print('저장', OUT)
