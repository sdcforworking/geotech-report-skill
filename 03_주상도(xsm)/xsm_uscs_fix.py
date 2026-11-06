"""실내시험 결과(SM)로 붕적층 3개 지층 고치기: No.359(B)·381(C)·382(D)
분류 'GM:실트질 자갈'→'SM', 지층모양 'GM'→'SM', 지층구성·설명 첫 줄 '모래섞인 실트질 자갈'→'자갈섞인 실트질 모래' (xsm에 이미 있는 표기).
설명 2번째 줄 이후·심도·색상·SPT·코어는 그대로. 다른 지층은 손대지 않음(전·후 파싱 비교).
사용: python xsm_uscs_fix.py 입력.xsm 출력.xsm"""
import sys, os, re, bisect, struct, json
OLD = r'C:\Users\심대천\AppData\Local\Temp\claude\C--Users-----Desktop-17-------\d4b9da42-737c-4228-af0b-c59dd9861c81\scratchpad'
here = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, OLD); sys.path.insert(1, here)
from xsm_tree import parse_file
from layer_names import layer_blocks
from layer_fields import layers
from xsm_patch import apply_edits, lstring
from xsmparse import strings, parse

SRC, DST = sys.argv[1:3]
HOLES = ('No.359(B)', 'No.381(C)', 'No.382(D)')
OLDC, NEWC = '모래섞인 실트질 자갈', '자갈섞인 실트질 모래'
b = open(SRC, 'rb').read()
root = parse_file(b)
HN = [(o, t) for o, e, t in strings(b) if re.fullmatch(r'No\.[\d\-]+\([A-Z\-]+\)', t)]
offs = [o for o, _ in HN]
owner = lambda p: HN[bisect.bisect(offs, p) - 1][1]


def desc_slot(b, x):          # first_line_sync.py와 같음
    p = x['f0_pos']
    for _ in range(4):
        p += 4 + struct.unpack_from('<I', b, p)[0]
    L = struct.unpack_from('<I', b, p)[0]
    return p, L, b[p + 4:p + 4 + L].decode('cp949')


LB = {x['block']: x for x in layer_blocks(b, root) if not x['title'].startswith('No.')}
edits = []
for f in layers(b, root):
    h = owner(f['block'])
    if h not in HOLES or f['name'] != '붕적층':
        continue
    x = LB[f['block']]
    assert (f['class'], f['shape'], x['name2']) == ('GM:실트질 자갈', 'GM', OLDC), (h, f['class'], f['shape'], x['name2'])
    dp, dL, d = desc_slot(b, x)
    assert d == f['desc'] and d.split('\r\n')[0] == '-' + OLDC, (h, d[:40])
    edits += [(f['class_pos'], 4 + f['class_len'], lstring('SM')),
              (f['shape_pos'], 4 + f['shape_len'], lstring('SM')),
              (dp, 4 + dL, lstring('\r\n'.join(['-' + NEWC] + d.split('\r\n')[1:]))),
              (x['name2_pos'], 4 + x['name2_len'], lstring(NEWC))]
    print(h, '붕적층 0~', x['bottom'], 'm 고침')
assert len(edits) == 12
nb, nsz = apply_edits(b, root, sorted(edits))
open(DST, 'wb').write(nb)

# 검증: 전·후 파싱 비교 — 세 지층의 desc 첫 줄·comp만 다르고 나머지(심도·SPT·코어·다른 지층) 같음
A, B = parse(SRC), parse(DST)
assert list(A) == list(B)
diff = []
for h in A:
    assert A[h]['spt'] == B[h]['spt'] and A[h]['core'] == B[h]['core'], h
    for la, lb in zip(A[h]['lay'], B[h]['lay']):
        for k in la:
            if la[k] != lb[k]:
                diff.append((h, la['ly'], k, la[k][:30], lb[k][:30]))
for e in diff:
    print('  ', e)
assert len(diff) == 6 and all(e[0] in HOLES and e[2] in ('desc', 'comp') for e in diff)
r2 = parse_file(nb)
L2 = [f for f in layers(nb, r2) if owner(f['block']) in HOLES and f['name'] == '붕적층']
print([(f['class'], f['shape']) for f in L2], '| size 보정', nsz, '| 다시 읽기 OK')
