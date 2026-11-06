"""풍화토·풍화암 판정 기준(50/10)과 다른 SPT 값, 공내재하시험 표의 지층·N값·암종을 4.3.2 공별 표와 대조.
사용: python check_more.py body.txt (hwpx_dump 결과)"""
import os, re, sys
sys.stdout.reconfigure(encoding='utf-8')
exec(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'check_spt.py'), encoding='utf-8').read().split('# 공별 지층표 N과 대조')[0])
# 풍화토인데 50/10 이하, 풍화암인데 50/11 이상
for h, vals in spt.items():
    for dep, v in vals:
        lay = next((x for x in layers[h] if x['f'] - 1e-6 <= dep < x['t'] - 1e-6), None) or next((x for x in layers[h] if abs(dep - x['t']) < 1e-6), None)
        a, b = map(int, v.split('/'))
        if lay['layer'] == '풍화토' and a == 50 and b <= 10:
            print('풍화토인데 50/10 이하', h, dep, v)
        if lay['layer'] == '풍화암' and (a < 50 or b > 10):
            print('풍화암인데 50/10 미달', h, dep, v)
        if lay['layer'] == '붕적층' and a == 50 and b <= 10:
            print('붕적층 50/10 이하', h, dep, v, lay['comp'])
# 공내재하 표 N값 대조
for t, d in tables.items():
    rows = d['rows']
    if rows and rows[0][0] == '철탑번호' and '변형계수' in ' '.join(rows[0]):
        for r in rows[1:]:
            h, dep, lay, rock, n = r[1], float(r[2]), r[3], r[4], r[5]
            L2 = layers[h]
            x = next((x for x in L2 if x['f'] - 1e-6 <= dep < x['t'] - 1e-6), None) or next((x for x in L2 if abs(dep - x['t']) < 1e-6), None)
            if x is None or x['layer'] != lay:
                print('공내재하 지층', h, dep, lay, '→ 4.3.2', x and x['layer'])
            if lay in ('풍화토', '풍화암'):
                sv = dict(spt.get(h, []))
                if sv.get(dep, '-') != n:
                    print('공내재하 N', h, dep, n, '→ SPT', sv.get(dep, '-'))
            else:
                if x and x['n'] != n:
                    print('공내재하 TCR/RQD', h, n, '→ 4.3.2', x['n'])
            if x and x['layer'] == '연암' and not x['comp'].startswith(rock + '의'):
                print('공내재하 암종', h, rock, '→ 4.3.2', x['comp'])
print('끝')
