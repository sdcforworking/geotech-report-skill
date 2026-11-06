"""일러스트 AI11 텍스트 문서(PostScript 비슷한 사전 형식) 읽기 — 값마다 원본 바이트 위치(시작·끝)를 같이 기억함."""
import re

TOK = re.compile(rb'\s*(<<|>>|\[|\]|/[^\s/\[\]<>()]*|\((?:[^\\)]|\\.)*\)|<[0-9A-Fa-f\s]*>|[^\s/\[\]<>()]+)', re.S)


class Str(bytes):
    """(…) 문자열. text 속성은 UTF-16이면 풀어 둔 값."""
    def __new__(cls, raw, start, end):
        o = super().__new__(cls, raw)
        o.start, o.end = start, end
        inner = re.sub(rb'\\(.)', rb'\1', raw[1:-1], flags=re.S)
        o.text = inner[2:].decode('utf-16-be', 'replace') if inner.startswith(b'\xfe\xff') else inner.decode('latin-1')
        return o


class Node:
    def __init__(self, kind, start):
        self.kind, self.start, self.end, self.items = kind, start, None, []

    def dict(self):
        it = self.items
        return {it[i]: it[i + 1] for i in range(0, len(it) - 1, 2)}


def parse(data: bytes):
    pos, stack = 0, [Node('root', 0)]
    n = len(data)
    while pos < n:
        m = TOK.match(data, pos)
        if not m:
            if data[pos:].strip() == b'':
                break
            raise ValueError(f'토큰 못 읽음 @{pos}: {data[pos:pos + 40]!r}')
        t, s, pos = m.group(1), m.start(1), m.end()
        if t in (b'<<', b'['):
            nd = Node('d' if t == b'<<' else 'a', s); stack[-1].items.append(nd); stack.append(nd)
        elif t in (b'>>', b']'):
            nd = stack.pop(); nd.end = pos
        elif t.startswith(b'('):
            stack[-1].items.append(Str(t, s, pos))
        else:
            stack[-1].items.append(t.decode('latin-1'))
    return stack[0]
