"""hwpx 본문을 검토용 텍스트로 풀어 씀 — 문단은 한 줄씩, 표는 칸 주소(행,열)대로 격자로, 수식은 스크립트, 그림은 [그림].
병합 칸은 첫 칸에만 글자를 두고 나머지는 빈칸. 머리말·꼬리말·각주·글상자 글자도 표시.
사용: python hwpx_dump.py 파일.hwpx 출력.txt"""
import sys, zipfile, re
import xml.etree.ElementTree as ET

HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
SRC, OUT = sys.argv[1], sys.argv[2]
SPECIAL = {'tab': '\t', 'lineBreak': '⏎', 'nbSpace': ' ', 'fwSpace': '　', 'hyphen': '-'}


def t_text(t):
    s = t.text or ''
    for ch in t:
        tag = ch.tag.replace(HP, '')
        s += SPECIAL.get(tag, '')
        s += ch.tail or ''
    return s


class Dumper:
    def __init__(self):
        self.lines = []
        self.ntbl = 0
        self.neq = 0

    def para_inline(self, p):
        """문단의 글자만(표·그림은 따로 처리하며 자리표시)."""
        out = []
        for run in p.findall(HP + 'run'):
            for el in run:
                tag = el.tag.replace(HP, '')
                if tag == 't':
                    out.append(t_text(el))
                elif tag == 'equation':
                    sc = el.find(HP + 'script')
                    self.neq += 1
                    out.append('〔수식: ' + ((sc.text or '').strip() if sc is not None else '') + '〕')
                elif tag == 'tbl':
                    out.append('〔표〕')
                elif tag == 'pic':
                    out.append('〔그림〕')
                elif tag in ('rect', 'ellipse', 'polygon', 'container', 'line', 'curve', 'arc', 'textart'):
                    txt = self.draw_text(el)
                    out.append('〔도형' + (': ' + txt if txt else '') + '〕')
                elif tag == 'ctrl':
                    for c in el:
                        ct = c.tag.replace(HP, '')
                        if ct in ('header', 'footer', 'footNote', 'endNote'):
                            sub = ' / '.join(self.sub_text(c))
                            out.append(f'〔{ct}: {sub}〕')
        return ''.join(out)

    def draw_text(self, el):
        res = []
        for dt in el.iter(HP + 'drawText'):
            for sl in dt.findall(HP + 'subList'):
                res += self.sub_text(sl)
        return ' / '.join(x for x in res if x.strip())

    def sub_text(self, el):
        res = []
        for p in el.iter(HP + 'p'):
            # 중첩 표 안 문단까지 iter로 다 잡히므로 여기선 간단히 글자만
            s = ''.join(t_text(t) for t in p.iter(HP + 't'))
            if s.strip():
                res.append(s)
        return res

    def cell_text(self, tc):
        sl = tc.find(HP + 'subList')
        parts = []
        if sl is None:
            return ''
        for p in sl.findall(HP + 'p'):
            s = self.para_inline(p)
            # 셀 안의 표(중첩)는 따로 풀어 줌
            for run in p.findall(HP + 'run'):
                for el in run:
                    if el.tag == HP + 'tbl':
                        s += ' {중첩표: ' + ' ‖ '.join(' | '.join(r) for r in self.table_grid(el)) + '}'
            parts.append(s)
        return ' ⏎ '.join(x for x in parts if x.strip()) if len(parts) > 1 else (parts[0] if parts else '')

    def table_grid(self, tbl):
        rows = int(tbl.get('rowCnt', 0)); cols = int(tbl.get('colCnt', 0))
        grid = [['' for _ in range(cols)] for _ in range(rows)]
        for tr in tbl.findall(HP + 'tr'):
            for tc in tr.findall(HP + 'tc'):
                a = tc.find(HP + 'cellAddr'); sp = tc.find(HP + 'cellSpan')
                r, c = int(a.get('rowAddr')), int(a.get('colAddr'))
                rs, cs = int(sp.get('rowSpan', 1)), int(sp.get('colSpan', 1))
                txt = self.cell_text(tc).replace('\n', ' ')
                if r < rows and c < cols:
                    mark = '' if rs == 1 and cs == 1 else f'⟨{rs}×{cs}⟩'
                    grid[r][c] = txt + mark
        return grid

    def caption(self, el):
        cap = el.find(HP + 'caption')
        if cap is None:
            return ''
        return ' / '.join(self.sub_text(cap))

    def walk_para(self, p, where):
        line = self.para_inline(p)
        if line.strip():
            self.lines.append(f'{where}\t{line}')
        for run in p.findall(HP + 'run'):
            for el in run:
                tag = el.tag.replace(HP, '')
                if tag == 'tbl':
                    self.ntbl += 1
                    cap = self.caption(el)
                    g = self.table_grid(el)
                    self.lines.append(f'{where}\t┏ 표{self.ntbl} ({el.get("rowCnt")}×{el.get("colCnt")})' + (f' 캡션: {cap}' if cap else ''))
                    for i, row in enumerate(g):
                        self.lines.append(f'{where}\t┃ r{i}: ' + ' | '.join(row))
                    self.lines.append(f'{where}\t┗')
                elif tag == 'pic':
                    cap = self.caption(el)
                    if cap:
                        self.lines.append(f'{where}\t〔그림 캡션: {cap}〕')


z = zipfile.ZipFile(SRC)
names = sorted([n for n in z.namelist() if re.match(r'Contents/section\d+\.xml', n)], key=lambda n: int(re.findall(r'\d+', n)[0]))
d = Dumper()
for n in names:
    root = ET.fromstring(z.read(n))
    sec = re.findall(r'\d+', n)[0]
    d.lines.append(f'===== section{sec} =====')
    for i, p in enumerate(root.findall(HP + 'p')):
        d.walk_para(p, f's{sec}p{i}')
with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(d.lines))
print('sections', names, '| lines', len(d.lines), '| tables', d.ntbl, '| equations', d.neq)
