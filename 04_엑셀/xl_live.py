"""열려 있는 엑셀 통합문서들의 '암종 구분' 시트 값을 읽기만 한다 (저장·수정 없음)."""
import json, win32com.client as w
xl = w.GetActiveObject('Excel.Application')
out = {}
for wb in xl.Workbooks:
    ws = wb.Worksheets('암종 구분')
    out[wb.FullName] = [list(r) for r in ws.Range('A1:H130').Value]
json.dump(out, open('xl_live_rock.json', 'w', encoding='utf-8'), ensure_ascii=False, default=str)
(a, A), (b, B) = out.items()
print('다른 행:', [(i, x, y) for i, (x, y) in enumerate(zip(A, B)) if x != y])
from collections import Counter
print(Counter(r[6] for r in A[1:123]))
for r in A[1:123]:
    if r[2] == '풍화암': print(r[1], '|', r[5], '|', r[6])
