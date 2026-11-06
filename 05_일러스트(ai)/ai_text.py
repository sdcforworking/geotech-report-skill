"""일러스트레이터 없이 .ai(PDF 호환, Illustrator 24 이상 zstd 압축) 안의 글상자 글자를 보거나 바꿈.

  보기:  python ai_text.py list 파일.ai [찾을글자]
         → 글상자(story) 번호와 글자, 틀 위치(tx, ty)를 보여 줌. 같은 글자가 여러 개면 위치로 고름.
  바꾸기: python ai_text.py replace 파일.ai 새파일.ai 글상자번호 "옛글자=새글자" ["옛=새" ...] [--font C:\\Windows\\Fonts\\malgunbd.ttf]
         → 그 글상자 안에서만 바꿈. 옛글자·새글자는 글자 수가 같아야 함(레이아웃 캐시를 그대로 쓰기 때문).
           --font를 주면 글상자의 글리프 번호 캐시(/5 [ … ])도 같은 글꼴 기준으로 같이 바꿈(권장).

원리: PieceInfo → Private 사전의 AIPrivateData1..N(64KB 조각) + /NumBlock 이 일러스트 원본.
앞에 '%AI24_ZStandard_Data'가 붙은 zstd 압축 → 풀면 PostScript 비슷한 글자 자료 →
'%AI11_BeginTextDocument' 안 ASCII85 텍스트 문서에 글자가 UTF-16BE로 들어 있음.
같은 길이로 바꾸면 ASCII85도 같은 길이로 바뀌어 다른 부분은 그대로 남음.
주의: PDF 미리보기(쪽 내용)는 바꾸지 않음 → 다른 뷰어·썸네일에는 옛 글자가 보이고, 일러스트레이터에서 열면 새 글자가 보임.
      미리보기까지 바꾸는 예는 ai_replace_example.py 참고(글꼴 조각에 없는 글리프 채워 넣기 포함).
항상 사본으로 작업하고 원본은 백업해 둘 것."""
import os, sys, re, base64
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fitz, zstd_ct, tdparse

sys.stdout.reconfigure(encoding='utf-8')
HDR = b'/AI11TextDocument : /ASCII85Decode ,'
ZHDR = b'%AI24_ZStandard_Data'
CHUNK = 65536


def find_private(doc):
    for i in range(doc.page_count):
        v = doc.xref_get_key(doc[i].xref, 'PieceInfo')
        m = re.search(r'/Illustrator (\d+) 0 R', v[1]) if v[0] == 'dict' else None
        if m:
            p = re.search(r'/Private (\d+) 0 R', doc.xref_object(int(m.group(1))))
            return int(p.group(1))
    raise RuntimeError('일러스트 원본 자료(PieceInfo) 없음')


def load(path):
    doc = fitz.open(path)
    px = find_private(doc)
    keys = {int(k): int(x) for k, x in re.findall(r'/AIPrivateData(\d+) (\d+) 0 R', doc.xref_object(px))}
    raw = b''.join(doc.xref_stream_raw(keys[i]) for i in range(1, len(keys) + 1))
    if not raw.startswith(ZHDR):
        raise RuntimeError('zstd 형식이 아님(오래된 버전 ai일 수 있음): ' + repr(raw[:24]))
    nat, rest = zstd_ct.decompress(raw[len(ZHDR):])
    return doc, px, keys, bytearray(nat)


def text_doc(nat):
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
    return pos, data, base64.a85decode(data)


def stories(dec):
    root = tdparse.parse(dec); root.kind = 'd'
    D = root.dict()
    sts = D['/1'].dict()['/1'].items
    frames = D['/0'].dict()['/8'].dict()['/0'].items
    out = []
    for i, st in enumerate(sts):
        sd = st.dict(); t = sd['/0'].dict()['/0']
        try:
            fi = int(sd['/1'].dict()['/0'].items[0].dict()['/0'])
            m = frames[fi].dict()['/0'].dict()['/2'].dict()['/2'].items
            tx, ty = float(m[4]), float(m[5])
        except Exception:
            fi, tx, ty = None, None, None
        out.append(dict(i=i, text=t.text if hasattr(t, 'text') else '', start=st.start, end=st.end, frame=fi, tx=tx, ty=ty))
    return out


def glyph_ids(font, s):
    return [font.has_glyph(ord(ch)) for ch in s]


def cmd_list(path, key=None):
    _, _, _, nat = load(path)
    _, _, dec = text_doc(nat)
    for s in stories(dec):
        if key is None or key in s['text']:
            print(f"{s['i']:4d}  틀 {s['frame']}  tx {s['tx']}  ty {s['ty']}  {s['text']!r}")


def cmd_replace(src, dst, idx, pairs, fontfile=None):
    doc, px, keys, nat = load(src)
    nat0 = bytes(nat)
    pos, data, dec0 = text_doc(nat)
    st = stories(dec0)[idx]
    print('바꿀 글상자', idx, repr(st['text']))
    dec = bytearray(dec0); edits = []

    def rep(old, new):
        i = dec0.find(old, st['start'], st['end'])
        if i < 0 or dec0.find(old, i + 1, st['end']) != -1:
            raise SystemExit(f'글상자 안에서 한 번만 나와야 함: {old!r}')
        dec[i:i + len(old)] = new; edits.append((i, len(old)))

    font = fitz.Font(fontfile=fontfile) if fontfile else None
    for o, n in pairs:
        if len(o) != len(n):
            raise SystemExit(f'글자 수가 달라요: {o} / {n}')
        rep(o.encode('utf-16-be'), n.encode('utf-16-be'))
        if font:   # 글리프 캐시: 바뀐 글자가 들어 있는 줄의 /5 [ … ]
            seg = dec0[st['start']:st['end']]
            for m in re.finditer(rb'/5 \[ ([0-9 ]+) \]', seg):
                ids = [int(x) for x in m.group(1).split()]
                oi, ni = glyph_ids(font, o), glyph_ids(font, n)
                for k in range(len(ids) - len(oi) + 1):
                    if ids[k:k + len(oi)] == oi:
                        new_ids = ids[:k] + ni + ids[k + len(oi):]
                        old_b = m.group(0); new_b = b'/5 [ ' + ' '.join(map(str, new_ids)).encode() + b' ]'
                        if len(old_b) != len(new_b):
                            raise SystemExit('글리프 번호 자릿수가 달라 같은 길이로 못 바꿈 — --font 없이 다시 해 보세요')
                        rep(old_b, new_b)
                        break
    groups, k = [], 0
    while k < len(pos):
        if data[k:k + 1] == b'z':
            groups.append((k, 1)); k += 1
        else:
            groups.append((k, min(5, len(pos) - k))); k += 5
    for i, n in edits:
        for g in range(i // 4, (i + n - 1) // 4 + 1):
            k0, ln = groups[g]
            if ln != 5:
                raise SystemExit('ASCII85 압축 기호(z) 칸에 걸림 — 예상 밖')
            enc = base64.a85encode(bytes(dec[4 * g:4 * g + 4]))
            for t in range(5):
                nat[pos[k0 + t]] = enc[t]
    assert base64.a85decode(bytes(nat[p] for p in pos)) == bytes(dec)
    changed = sum(1 for x, y in zip(nat, nat0) if x != y)
    comp = ZHDR + zstd_ct.compress(bytes(nat), 3)
    chunks = [comp[i:i + CHUNK] for i in range(0, len(comp), CHUNK)]
    for i, ch in enumerate(chunks, 1):
        if i in keys:
            doc.update_stream(keys[i], ch, compress=False)
        else:
            x = doc.get_new_xref(); doc.update_object(x, '<<>>'); doc.update_stream(x, ch, new=True, compress=False)
            doc.xref_set_key(px, f'AIPrivateData{i}', f'{x} 0 R')
    for i in range(len(chunks) + 1, len(keys) + 1):
        doc.xref_set_key(px, f'AIPrivateData{i}', 'null')
    if doc.xref_get_key(px, 'NumBlock')[0] != 'null':
        doc.xref_set_key(px, 'NumBlock', str(len(chunks)))
    doc.save(dst, garbage=1)
    print(f'바뀐 바이트 {changed} | 조각 {len(keys)} → {len(chunks)} | 저장 {dst}')
    print('※ PDF 미리보기는 옛 글자 그대로 — 일러스트레이터에서 열어 확인하세요')


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a or a[0] not in ('list', 'replace'):
        sys.exit(__doc__)
    if a[0] == 'list':
        cmd_list(a[1], a[2] if len(a) > 2 else None)
    else:
        font = None
        if '--font' in a:
            j = a.index('--font'); font = a[j + 1]; a = a[:j] + a[j + 2:]
        pairs = [tuple(p.split('=', 1)) for p in a[4:]]
        cmd_replace(a[1], a[2], int(a[3]), pairs, font)
