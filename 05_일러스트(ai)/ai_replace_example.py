"""지형및지질 ai 사본에서 산계·수계도(2번째 대지) 오른쪽 위 '벽학산 (331.1m)' 글상자를 '오음산 (929.4m)'으로 바꿈.
① 일러스트 원본 데이터(AIPrivateData, zstd): 텍스트 문서의 해당 글상자(story 434) 글자와 글리프 번호를 같은 길이로 바꾼 뒤 다시 압축
② PDF 미리보기: 글꼴 조각에 '음' 글리프를 맑은고딕 굵게에서 채워 넣고, 해당 글자 그리기와 흰 테두리를 새 글자로 바꿈
아래쪽 '벽학산 (331.1m)'(story 433)은 건드리지 않음.
사용: python make_new_ai.py 원본사본.ai 새파일.ai"""
import sys, re, struct, base64
sys.path.insert(0, '.')
sys.stdout.reconfigure(encoding='utf-8')
import fitz, zstd_ct, tdparse

SRC, DST = sys.argv[1], sys.argv[2]
HDR = b'/AI11TextDocument : /ASCII85Decode ,'
ZHDR = b'%AI24_ZStandard_Data'
CHUNK = 65536

doc = fitz.open(SRC)

# ---------- ① 원본 데이터 ----------
priv = doc.xref_object(25)
keys = dict((int(k), int(x)) for k, x in re.findall(r'/AIPrivateData(\d+) (\d+) 0 R', priv))
others = [k for k in re.findall(r'/(\w+)', priv) if not k.startswith('AIPrivateData')]
print('Private 사전 다른 키:', others)
assert set(others) <= {'AIMetaData', 'ContainerVersion', 'CreatorVersion', 'NumBlock', 'RoundtripStreamType', 'RoundtripVersion'}, others
old_n = len(keys)
assert doc.xref_get_key(25, 'NumBlock') == ('int', str(old_n))
raw = b''.join(doc.xref_stream_raw(keys[i]) for i in range(1, old_n + 1))
assert raw.startswith(ZHDR)
nat0, rest = zstd_ct.decompress(raw[len(ZHDR):]); assert rest == 0
nat = bytearray(nat0)

a = nat.index(HDR) + len(HDR); end = nat.index(b'~>', a)
pos, line_start = [], False
for j in range(a, end):
    c = nat[j]
    if c in (13, 10):
        line_start = c == 10; continue
    if line_start and c == ord('%'):
        line_start = False; continue
    line_start = False
    pos.append(j)
data = bytes(nat[p] for p in pos)
dec0 = base64.a85decode(data)
groups, k = [], 0
while k < len(pos):
    if data[k:k + 1] == b'z':
        groups.append((k, 1)); k += 1
    else:
        groups.append((k, min(5, len(pos) - k))); k += 5

root = tdparse.parse(dec0); root.kind = 'd'
stories = root.dict()['/1'].dict()['/1'].items
st_new, st_keep = stories[434], stories[433]
assert st_new.dict()['/0'].dict()['/0'].text == '벽학산\r(331.1m)\r' == st_keep.dict()['/0'].dict()['/0'].text
assert st_new.dict()['/1'].dict()['/0'].items[0].dict()['/0'] == '371'      # 위쪽(틀 371, ty -184.5)
s0, s1 = st_new.start, st_new.end
dec = bytearray(dec0)
edits = []


def rep(old, new):
    i = dec0.find(old, s0, s1)
    assert i >= 0 and dec0.find(old, i + 1, s1) == -1 and len(old) == len(new), old
    dec[i:i + len(old)] = new; edits.append((i, len(old)))


rep('벽학'.encode('utf-16-be'), '오음'.encode('utf-16-be'))
rep('331.1'.encode('utf-16-be'), '929.4'.encode('utf-16-be'))
rep(b'/5 [ 1586 2842 1740 3 ]', b'/5 [ 2074 2165 1740 3 ]')
rep(b'/5 [ 10 21 21 19 16 19 79 11 3 ]', b'/5 [ 10 27 20 27 16 22 79 11 3 ]')
for i, n in edits:
    for g in range(i // 4, (i + n - 1) // 4 + 1):
        k0, ln = groups[g]
        assert ln == 5, '압축 기호(z) 칸은 예상 밖'
        enc = base64.a85encode(bytes(dec[4 * g:4 * g + 4]))
        assert len(enc) == 5
        for t in range(5):
            nat[pos[k0 + t]] = enc[t]
data2 = bytes(nat[p] for p in pos)
assert base64.a85decode(data2) == bytes(dec), '다시 풀어 본 텍스트 문서가 다름'
assert len(nat) == len(nat0)
open('native_new.bin', 'wb').write(nat)
comp = ZHDR + zstd_ct.compress(bytes(nat), 3)
chunks = [comp[i:i + CHUNK] for i in range(0, len(comp), CHUNK)]
print(f'원본 데이터: 압축 {len(raw)} → {len(comp)}, 조각 {old_n} → {len(chunks)}')
for i, ch in enumerate(chunks, 1):
    if i in keys:
        doc.update_stream(keys[i], ch, compress=False)
    else:
        x = doc.get_new_xref(); doc.update_object(x, '<<>>'); doc.update_stream(x, ch, new=True, compress=False)
        doc.xref_set_key(25, f'AIPrivateData{i}', f'{x} 0 R')
for i in range(len(chunks) + 1, old_n + 1):
    doc.xref_set_key(25, f'AIPrivateData{i}', 'null')
doc.xref_set_key(25, 'NumBlock', str(len(chunks)))

# ---------- ② PDF 미리보기 ----------
# 글꼴 조각(C2_0, FontFile2 2766)에 '음'(GID 2165) 채우기
G = 2165
ff = bytearray(doc.xref_stream(2766))


def tabdir(b):
    n = struct.unpack('>H', b[4:6])[0]
    return n, [list(struct.unpack('>4sIII', b[12 + 16 * i:28 + 16 * i])) for i in range(n)]


mg = open(r'C:\Windows\Fonts\malgunbd.ttf', 'rb').read()
_, mt = tabdir(mg); mt = {t[0]: t for t in mt}
assert struct.unpack('>h', mg[mt[b'head'][2] + 50:mt[b'head'][2] + 52])[0] == 1
ml, mgl = mt[b'loca'][2], mt[b'glyf'][2]
ga, gb = struct.unpack('>II', mg[ml + 4 * G:ml + 4 * G + 8])
gdata = mg[mgl + ga:mgl + gb]
assert struct.unpack('>h', gdata[:2])[0] > 0          # 단순 글리프
gdata += b'\0' * (-len(gdata) % 4)
n, td = tabdir(ff); T = {t[0]: t for t in td}
assert struct.unpack('>h', ff[T[b'head'][2] + 50:T[b'head'][2] + 52])[0] == 1
loca = bytearray(ff[T[b'loca'][2]:T[b'loca'][2] + T[b'loca'][3]])
glyf = bytearray(ff[T[b'glyf'][2]:T[b'glyf'][2] + T[b'glyf'][3]])
la, lb = struct.unpack('>II', loca[4 * G:4 * G + 8]); assert la == lb, '이미 글리프가 있음'
glyf[la:la] = gdata
ng = len(loca) // 4
for g in range(G + 1, ng):
    v = struct.unpack('>I', loca[4 * g:4 * g + 4])[0]
    loca[4 * g:4 * g + 4] = struct.pack('>I', v + len(gdata))
newtab = {b'loca': bytes(loca), b'glyf': bytes(glyf)}
order = sorted(td, key=lambda t: t[2])                 # 원래 파일 안 순서대로 다시 씀
body, off, rec = b'', 12 + 16 * n, {}
for t in order:
    tb = newtab.get(t[0], bytes(ff[t[2]:t[2] + t[3]]))
    if t[0] == b'head':
        tb = tb[:8] + b'\0\0\0\0' + tb[12:]
    cs = sum(struct.unpack('>%dI' % ((len(tb) + 3) // 4), tb + b'\0' * (-len(tb) % 4))) & 0xFFFFFFFF
    rec[t[0]] = (cs, off, len(tb))
    tb += b'\0' * (-len(tb) % 4); body += tb; off += len(tb)
dirb = bytes(ff[:12]) + b''.join(struct.pack('>4sIII', t[0], *rec[t[0]]) for t in td)
font = bytearray(dirb + body)
adj = (0xB1B0AFBA - sum(struct.unpack('>%dI' % (len(font) // 4), bytes(font))) ) & 0xFFFFFFFF
ho = rec[b'head'][1]; font[ho + 8:ho + 12] = struct.pack('>I', adj)
doc.update_stream(2766, bytes(font)); doc.xref_set_key(2766, 'Length1', str(len(font)))
cidset = bytearray(doc.xref_stream(2765))
if len(cidset) > G // 8:
    cidset[G // 8] |= 0x80 >> (G % 8); doc.update_stream(2765, bytes(cidset))
print('글꼴 조각:', len(ff), '→', len(font), '| 음 글리프', len(gdata), '바이트')

# 쪽 내용: 위쪽 글상자 두 번의 글자 그리기 + 그 사이 흰 테두리
cs = doc.xref_stream(2239)
TM = b'7.6 0 0 8 378.2627 241.4536 Tm'
occ = [m.start() for m in re.finditer(re.escape(TM), cs)]
assert len(occ) == 2
b1 = cs.rfind(b'BT', 0, occ[0]); e1 = cs.find(b'ET', occ[0]) + 2
b2 = cs.rfind(b'BT', 0, occ[1]); e2 = cs.find(b'ET', occ[1]) + 2
halo = cs[e1:b2]
assert halo.count(b' cm\n') == 11 and halo.strip().startswith(b'q') and halo.strip().endswith(b'Q')
hp = halo.index(b'1 w 4 M') + len(b'1 w 4 M')
OLD_K, NEW_K = b'<06320B1A06CC>', b'<081A087506CC>'
OLD_N, NEW_N = b'(\\(331.1m\\))', b'(\\(929.4m\\))'
t1 = cs[b1:e1].replace(OLD_K, NEW_K).replace(OLD_N, NEW_N)
t2 = cs[b2:e2].replace(OLD_K, NEW_K).replace(OLD_N, NEW_N)
assert NEW_K in t1 and NEW_N in t1 and NEW_K in t2 and NEW_N in t2
stroke = (b' \nBT\n1 Tr\n/C2_0 1 Tf\n' + TM + b'\n' + NEW_K + b'Tj\n/TT1 1 Tf\n-0.485 -1.2 Td\n' + NEW_N + b'Tj\nET\nQ\n')
new_halo = halo[:hp] + stroke
cs2 = cs[:b1] + t1 + new_halo + t2 + cs[e2:]
doc.update_stream(2239, cs2)
print('쪽 내용:', len(cs), '→', len(cs2))

doc.save(DST, garbage=1)
print('저장', DST)
