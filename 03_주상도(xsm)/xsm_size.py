"""XsiteMap xsm 파일에서 시추공별 시추 규격(BX/NX 등)을 뽑음.
xsm은 바이너리지만 공번(No.000(A) 꼴)과 규격 글자(BX·NX·HX·PX)가 cp949로 그대로 들어 있어, 규격 글자 앞에서 가장 가까운 공번을 짝지음.
공번 모양이 다른 과업이면 --hole 정규식을 바꿔서 씀.

사용: python xsm_size.py 통합.xsm [출력.json] [--hole "No\\.[0-9\\-]+\\([A-D\\-]+\\)"]"""
import argparse, json, re, sys
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')
ap = argparse.ArgumentParser()
ap.add_argument('xsm'); ap.add_argument('out', nargs='?')
ap.add_argument('--hole', default=r'No\.[0-9\-]+\([A-D\-]+\)')
ap.add_argument('--sizes', default='BX,NX,HX,PX,AX,NQ,HQ')
a = ap.parse_args()
b = open(a.xsm, 'rb').read()
hole = re.compile(a.hole.encode())
size = re.compile(b'(' + b'|'.join(x.encode() for x in a.sizes.split(',')) + b')')
out = {}
for m in size.finditer(b):
    h = hole.findall(b[max(0, m.start() - 400):m.start()])
    if h:
        out[h[-1].decode('cp949')] = m.group(1).decode()   # 공 머리 자료 안 규격 칸(공번 바로 뒤)
print('공 수', len(out), dict(Counter(out.values())))
if a.out:
    json.dump(out, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('저장', a.out)
