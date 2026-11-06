"""XsiteMap 2.0 통합 xsm에서 공별 지층·SPT·코어 기록을 직접 추출."""
import sys, struct, re, json

LAY = ('붕적층', '풍화토', '풍화암', '연암')


def strings(d):
    out = []
    i, L = 0, len(d)
    while i < L - 4:
        n = d[i] | (d[i + 1] << 8)
        if 1 <= n <= 4000 and d[i + 2] == 0 and d[i + 3] == 0 and i + 4 + n <= L:
            s = d[i + 4:i + 4 + n]
            if all(b >= 32 or b in (9, 10, 13) for b in s):
                try:
                    out.append((i, i + 4 + n, s.decode('cp949')))
                    i += 4 + n
                    continue
                except UnicodeDecodeError:
                    pass
        i += 1
    return out


def nice(d, a, b):
    """a~b 구간에서 소수 둘째 자리 이하로 딱 떨어지는 첫 double."""
    for o in range(a, b - 7):
        v = struct.unpack_from('<d', d, o)[0]
        if 0.05 < v < 500 and abs(v * 100 - round(v * 100)) < 1e-7:
            return o, round(v, 2)
    return None, None


def parse(path):
    d = open(path, 'rb').read()
    S = strings(d)
    H = {}
    cur = None
    for k, (o, e, t) in enumerate(S):
        if re.fullmatch(r'No\.[\d\-]+\([A-Z\-]+\)', t) and k + 1 < len(S) and '시추' in S[k + 1][2]:
            cur = t
            H[cur] = {'lay': [], 'spt': [], 'core': []}
            continue
        if cur is None:
            continue
        # 지층: 이름 → '데이타' → … → 설명('-'로 시작) → 이름 → 색상 → 지층구성
        if t in LAY and S[k + 1][2] == '데이타':
            j = next(j for j in range(k + 2, k + 10) if S[j][2].startswith('-'))
            desc = S[j][2]
            assert S[j + 1][2] == t, (cur, t, S[j + 1])
            colf, comp = S[j + 2][2], S[j + 3][2]
            bot = round(struct.unpack_from('<d', d, S[j + 2][1])[0], 2)
            H[cur]['lay'].append(dict(ly=t, desc=desc, colf=colf, comp=comp, bot=bot))
        # SPT: '  d.dd[...]' 시료 기록
        m = re.fullmatch(r'\s*([\d.]+)\[(.+)\]', t)
        if m and S[k + 1][2] == '데이타':
            dep = struct.unpack_from('<d', d, e + 284)[0]
            n = struct.unpack_from('<i', d, e + 292)[0]
            p = struct.unpack_from('<d', d, e + 296)[0]
            H[cur]['spt'].append(dict(dep=round(dep, 2), n=n, p=round(p, 1), kind=m.group(2), label=float(m.group(1))))
        if re.fullmatch(r'[\d.]+ ∼ [\d.]+', t.strip()):
            a, b = [float(x) for x in t.split('∼')]
            H[cur]['core'].append((a, b, [S[k + j][2] for j in range(3, 11)]))
    return H


if __name__ == '__main__':
    H = parse(sys.argv[1])
    json.dump(H, open(sys.argv[2], 'w', encoding='utf-8'), ensure_ascii=False)
    print(len(H), 'holes')
