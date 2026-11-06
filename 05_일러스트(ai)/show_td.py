"""textdoc.bin(일러스트 텍스트 문서) 특정 위치 앞뒤를 사람이 읽을 수 있게 보여 줌 — (þÿ…) 문자열은 UTF-16으로 풀어서 표시."""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
dec = open('textdoc.bin', 'rb').read()


def pretty(seg: bytes) -> str:
    def rep(m):
        raw = re.sub(rb'\\(.)', rb'\1', m.group(1), flags=re.S)
        return ('⟪' + raw.decode('utf-16-be', 'replace') + '⟫').encode('utf-8')
    out = re.sub(rb'\(\xfe\xff((?:[^\\)]|\\.)*)\)', rep, seg, flags=re.S)
    return out.decode('utf-8', 'replace')


for a in sys.argv[1:]:
    pos, b, f = (int(x) for x in a.split(':'))
    print('=====', pos)
    print(pretty(dec[pos - b:pos + f]))
