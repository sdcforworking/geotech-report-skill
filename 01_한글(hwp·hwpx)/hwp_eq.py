"""HWP(5.x) 본문에서 수식 개체(EQEDIT) 스크립트를 문단 글과 함께 순서대로 뽑음 (읽기 전용).
사용: python hwp_eq.py 파일.hwp 키워드시작 개수"""
import sys, zlib, struct, olefile
f, key, cnt = sys.argv[1], sys.argv[2], int(sys.argv[3])
ole = olefile.OleFileIO(f)
comp = ole.openstream('FileHeader').read()[36] & 1
secs = sorted([s for s in ole.listdir() if s[0] == 'BodyText'], key=lambda s: int(s[1][7:]))
items = []
for s in secs:
    d = ole.openstream(s).read()
    if comp: d = zlib.decompress(d, -15)
    p = 0
    while p < len(d):
        h = struct.unpack_from('<I', d, p)[0]; p += 4
        tag, size = h & 0x3FF, h >> 20
        if size == 0xFFF: size = struct.unpack_from('<I', d, p)[0]; p += 4
        b = d[p:p + size]; p += size
        if tag == 67:
            t = ''.join(ch for ch in b.decode('utf-16le', 'ignore') if ch >= ' ')
            if t.strip(): items.append(('T', t.strip()))
        elif tag == 88:
            n = struct.unpack_from('<H', b, 4)[0]
            items.append(('EQ', b[6:6 + 2 * n].decode('utf-16le', 'ignore')))
start = next(i for i, x in enumerate(items) if x[0] == 'T' and key in x[1])
for k, v in items[start:start + cnt]:
    print(('   [수식] ' if k == 'EQ' else '') + v[:160])
