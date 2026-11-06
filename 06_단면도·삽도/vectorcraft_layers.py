"""지층단면도 ai 전부를 레이어별로 나눠 사본 폴더에 저장하고, 정리 전후 그림이 같은지 확인.
원본 폴더는 읽기만 함. 결과: 출력폴더\\*.ai, 출력폴더\\_check\\*.png, 출력폴더\\_report.json
사용: python batch_layers.py 원본ai폴더 출력폴더"""
import json, os, re, subprocess, sys
from PIL import Image, ImageChops

sys.stdout.reconfigure(encoding='utf-8')
CLI = r'C:\Program Files\ArtCraft\VectorCraft\vectorcraft-cli.exe'
SRC, OUT = sys.argv[1], sys.argv[2]
CHK = os.path.join(OUT, '_check'); os.makedirs(CHK, exist_ok=True)
STRATA = {'표토층', '표토', '붕적층', '매립층', '퇴적층', '충적층', '풍화토', '풍화암', '연암', '보통암', '경암', '연암층', '경암층'}
ORDER = ['지층(색)', '지층선', '지반고', '표고축', '거리값', '시추주상', '프로파일', '지층명', '공번']


def run(args, steps=None):
    if steps is not None:
        sf = os.path.join(CHK, '_steps.json'); json.dump(steps, open(sf, 'w', encoding='utf-8'), ensure_ascii=False)
        args = args + ['--commands', sf]
    p = subprocess.run([CLI, 'run'] + args, capture_output=True)
    lines = [json.loads(l) for l in p.stdout.decode('utf-8', 'replace').splitlines() if l.strip().startswith('{')]
    errs = [d for d in lines if d.get('error')]
    if p.returncode or errs:
        raise RuntimeError((errs[0].get('error') if errs else p.stderr.decode('utf-8', 'replace'))[:300])
    return lines


def B(c):
    b = c['bounds']; return b['x'], b['y'], b['width'], b['height']


def classify(objs):
    cols = [B(c) for c in objs if c['kind'] == 'Path' and c['stroke'] == '#000000' and abs(float(c['strokeWidth']) - 0.4) < 0.01
            and c.get('anchors') == 4 and B(c)[3] > 10]
    if not cols:
        raise RuntimeError('시추 기둥을 못 찾음')
    left = min(x for x, y, w, h in cols); bottom = max(y + h for x, y, w, h in cols)

    def inside(c, pad=0.7):
        x, y, w, h = B(c)
        return any(x >= cx - pad and y >= cy - pad and x + w <= cx + cw + pad and y + h <= cy + chh + pad for cx, cy, cw, chh in cols)

    def near_col(c):
        x, y, w, h = B(c)
        # 기둥 옆 짧은 지시선: 가로선, 또는 N값 글자가 몰릴 때 비스듬히 벌어지는 짧은 선
        return (h < 0.6 or (w < 5 and c.get('anchors') == 2)) and w < 20 and any(abs(x + w - cx) < 1.5 or abs(x - (cx + cw)) < 1.5 for cx, cy, cw, chh in cols)

    def cls(c):
        k, n = c['kind'], (c.get('name') or '').replace(' ', '')
        x, y, w, h = B(c)
        if k == 'Type':
            if n.startswith('No.'): return '공번'
            if n in STRATA: return '지층명'
            if x + w < left - 15: return '표고축'
            if n.startswith('EL.(') or re.fullmatch(r'\d+/\d+', n) or re.fullmatch(r'\d+(\.\d+)?', n): return '프로파일'
            if re.fullmatch(r'[\d.]+m', n): return '거리값'
            return '기타'
        if c['stroke'] == '#e60012': return '공번'
        if c['fill'] not in ('None', '#000000', '#ffffff') and c['stroke'] == 'None': return '지층(색)'
        if x + w < left - 15: return '표고축'
        if y >= bottom + 2: return '거리값'
        if inside(c): return '시추주상'
        if near_col(c): return '프로파일'
        if w < 0.2 and c.get('anchors') == 2 and any(abs(y + h - cy) < 1 and cx <= x <= cx + cw for cx, cy, cw, chh in cols):
            return '공번'
        if c['stroke'] == '#000000' and c.get('anchors', 0) > 4: return '지층선'
        if c['stroke'] == '#000000' and c.get('anchors') == 2: return '지반고'
        return '기타'

    g = {}
    for c in objs:
        g.setdefault(cls(c), []).append(c)
    return g


report = []
files = sorted(f for f in os.listdir(SRC) if f.lower().endswith('.ai'))
for f in files:
    src = os.path.join(SRC, f); dst = os.path.join(OUT, f)
    base = os.path.splitext(f)[0]
    pb, pa = os.path.join(CHK, base + '_전.png'), os.path.join(CHK, base + '_후.png')
    item = {'file': f}
    try:
        ins = run(['--in', src, '--cmd', 'document.inspect', '--params', json.dumps({'depth': 2, 'childLimit': 100000})])
        layers = ins[-1]['result']['layers']
        objs = [c for L in layers for c in (L.get('children') or [])]
        if any((L.get('childCount') or 0) > len(L.get('children') or []) for L in layers):
            raise RuntimeError('객체 목록이 잘림')
        if any(c['kind'] in ('Group', 'Layer') for c in objs):
            raise RuntimeError('그룹·하위 레이어가 있어 이 방식 대상 아님')
        g = classify(objs)
        item['groups'] = {k: len(v) for k, v in g.items()}
        if '기타' in g:
            item['기타'] = [(c['kind'], c.get('name'), [round(v, 1) for v in B(c)]) for c in g['기타']]
            raise RuntimeError(f"분류 안 된 객체 {len(g['기타'])}개")
        steps = [{'command': 'document.export', 'params': {'path': pb, 'scale': 3}}]
        ref = {}
        for name in ORDER:
            if name in g:
                ref[name] = len(steps); steps.append({'command': 'layer.new', 'params': {'name': name}})
        for name in ORDER:
            if name in g:
                steps.append({'command': 'layer.move', 'params': {'ids': [c['id'] for c in g[name]], 'target': f'${ref[name]}.id'}})
        steps.append({'command': 'layer.delete', 'params': {'ids': [L['id'] for L in layers]}})
        steps.append({'command': 'file.save', 'params': {'path': dst}})
        run(['--in', src], steps)
        # 저장본을 다시 열어 그림·레이어 확인
        chk = run(['--in', dst, '--cmd', 'document.inspect', '--params', json.dumps({'depth': 1, 'childLimit': 0}), '--export', pa, '--scale', '3'])
        lay2 = [x for x in chk if x.get('result', {}).get('layers')][0]['result']['layers']
        item['layers'] = [(L['name'], L.get('childCount')) for L in lay2]
        if sum(n or 0 for _, n in item['layers']) != len(objs):
            raise RuntimeError('객체 수가 달라짐')
        a, b = Image.open(pb).convert('RGB'), Image.open(pa).convert('RGB')
        bb = ImageChops.difference(a, b).getbbox() if a.size == b.size else 'size'
        item['diff'] = bb
        item['ok'] = bb is None
    except Exception as e:
        item['ok'] = False; item['err'] = str(e)
    report.append(item)
    print(('OK ' if item['ok'] else 'XX ') + f, item.get('err', ''), item.get('diff', ''))
json.dump(report, open(os.path.join(OUT, '_report.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('완료', sum(r['ok'] for r in report), '/', len(report))
