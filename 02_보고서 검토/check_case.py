"""5.2 기존 설계사례: 개별 사례표의 최소~최대와 요약표를 칸마다 대조. 표 번호(253~268)는 과업마다 다르니 names·cases를 고쳐서 씀.
사용: python check_case.py body.txt (hwpx_dump 결과)"""
import os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'check_spt.py'), encoding='utf-8').read().split('def nval')[0])
names = {253: '와룡~법전(유사과업)', 254: '강릉~제진(유사과업)', 255: '동부3공구 No.14', 256: '동부1공구', 257: '동부3공구 No.83 등', 258: '동해안변환소', 259: '남사~진위',
         260: '강릉~제진(유사지반)', 261: '신분당선', 262: '군도3호선', 263: '오미재', 264: '남해 서면-여수', 265: '강릉~제진①(편마암)', 266: '강릉~제진②(편마암)'}
def cols(t):
    rows = tables[t]['rows']
    if rows[1][0] == '' or rows[1][0].startswith('단위') is False:
        # 두 줄 머리글
        top, sub = rows[0], rows[1]
        hdr, last = [], ''
        for a, b in zip(top[1:], sub[1:]):
            last = a or last
            hdr.append((last + ('(' + b + ')' if b else '')).replace('붕적층(', '붕적층(').strip())
        body_rows = rows[2:]
    else:
        hdr = rows[0][1:]; body_rows = rows[1:]
    out = {}
    for r in body_rows:
        par = r[0].split('(')[0]
        for h, v in zip(hdr, r[1:]):
            h2 = h.replace('매립층(모래)', '붕적층(모래)')
            h2 = re.sub(r'연암\(.*\)', '연암', h2)
            out[(h2, par)] = v
    return out
def rng(v):
    v = v.replace(',', '').replace('~', '∼')
    if v in ('-', ''): return None
    p = [float(x) for x in v.split('∼')]
    return (p[0], p[-1])
for summ, cases in ((267, [253, 254, 260, 261, 262, 263, 264]), (268, [255, 256, 257, 258, 259, 265, 266])):
    S = cols(summ)
    print('==', '화강암' if summ == 267 else '편마암', '요약표 대조 ==')
    C = {t: cols(t) for t in cases}
    for key, sv in S.items():
        vals = []
        for t, d in C.items():
            v = d.get(key)
            if key[0] == '붕적층(모래)' and v is None:
                v = d.get(('붕적층', key[1])) if t != 259 else None
            if v is None and key[0] == '붕적층(자갈)':
                v = None
            r = rng(v) if v else None
            if r: vals.append((t, r))
        if not vals: continue
        lo = min(x[1][0] for x in vals); hi = max(x[1][1] for x in vals)
        s = rng(sv)
        if s is None or abs(s[0] - lo) > 1e-6 or abs(s[1] - hi) > 1e-6:
            who_lo = [names[t] for t, r in vals if abs(r[0] - lo) < 1e-6]
            who_hi = [names[t] for t, r in vals if abs(r[1] - hi) < 1e-6]
            print(f'  {key[0]} {key[1]}: 요약 "{sv}" / 개별 사례 최소~최대 {lo:g}∼{hi:g} (최소 {who_lo}, 최대 {who_hi})')
