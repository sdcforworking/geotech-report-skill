import sys; sys.path.insert(0, '.')
sys.stdout.reconfigure(encoding='utf-8')
import tdparse
dec = open('textdoc.bin', 'rb').read()
root = tdparse.parse(dec); root.kind = 'd'


def desc(x, depth=0, maxd=3, maxk=14):
    pad = '  ' * depth
    if isinstance(x, tdparse.Node):
        if x.kind == 'd':
            dd = x.dict()
            print(pad + f'<< {len(dd)} keys @{x.start}-{x.end}')
            if depth < maxd:
                for k, v in list(dd.items())[:maxk]:
                    if isinstance(v, tdparse.Node):
                        print(pad + '  ' + str(k) + ':'); desc(v, depth + 2, maxd, maxk)
                    else:
                        print(pad + '  ' + str(k) + ': ' + (('⟪' + v.text[:40] + '⟫') if isinstance(v, tdparse.Str) else repr(v)[:60]))
        else:
            print(pad + f'[ {len(x.items)} items @{x.start}-{x.end}')
            if depth < maxd and x.items:
                desc(x.items[0], depth + 1, maxd, maxk)
    else:
        print(pad + repr(x)[:60])


if __name__ == '__main__':
    desc(root, 0, int(sys.argv[1]) if len(sys.argv) > 1 else 5)
