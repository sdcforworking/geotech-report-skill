"""HWP 본문에서 키워드 뒤 첫 표의 칸 주소·병합·글자를 읽음 (읽기 전용). 사용: python hwp_tbl.py 파일.hwp 키워드"""
import sys, zlib, struct, olefile
f, key = sys.argv[1], sys.argv[2]
nth = int(sys.argv[3]) if len(sys.argv) > 3 else 0
ole = olefile.OleFileIO(f)
comp = ole.openstream('FileHeader').read()[36] & 1
secs = sorted([s for s in ole.listdir() if s[0] == 'BodyText'], key=lambda s: int(s[1][7:]))
recs = []
for s in secs:
    d = ole.openstream(s).read()
    if comp: d = zlib.decompress(d, -15)
    p = 0
    while p < len(d):
        h = struct.unpack_from('<I', d, p)[0]; p += 4
        tag, lvl, size = h & 0x3FF, (h >> 10) & 0x3FF, h >> 20
        if size == 0xFFF: size = struct.unpack_from('<I', d, p)[0]; p += 4
        recs.append((tag, lvl, d[p:p + size])); p += size
txt = lambda b: ''.join(ch for ch in b.decode('utf-16le', 'ignore') if ch >= ' ').strip()
i0 = next(i for i, (t, l, b) in enumerate(recs) if t == 67 and key in txt(b))
it = [i for i in range(i0, len(recs)) if recs[i][0] == 77][nth]   # HWPTAG_TABLE
tl = recs[it][1]
cells, cur = [], None
for t, l, b in recs[it + 1:]:
    if l < tl: break
    if t == 72 and l == tl:   # LIST_HEADER = 칸
        col, row, cs, rs = struct.unpack_from('<4H', b, 8)
        cur = [row, col, rs, cs, '']; cells.append(cur)
    elif t == 67 and cur is not None:
        cur[4] += txt(b) + ' '
for c in cells: print(c)
