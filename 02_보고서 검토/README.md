# 02_보고서 검토

`01_한글/hwpx_dump.py`로 만든 body.txt를 읽어 보고서 안 표끼리 숫자가 맞는지 대조한다. 표 머리글(시추공번, 지 층, 철탑번호 등)과 절 제목(4.3.2, 4.4.1 …)으로 표를 찾으므로, 양식이 같은 과업이면 그대로 쓰고 다르면 찾는 글자를 고친다.

- `check_tables.py`: 공별 표 깊이·두께 연속성, 지층두께 총괄표, 지층특성 요약표, 지층개황 서술(층후·N값·상대밀도·TCR/RQD) 대조
- `check_spt.py`: SPT 표 횟수·합계, 공별 지층표 N값 범위와 SPT 값 대조 (`check_more.py`·`check_case.py`가 이 파일 앞부분을 불러 씀)

| 파일 | 종류 | 설명 |
|---|---|---|
| `check_case.py` | 바로 사용 | 5.2 기존 설계사례: 개별 사례표의 최소~최대와 요약표를 칸마다 대조. 표 번호(253~268)는 과업마다 다르니 names·cases를 고쳐서 씀. 사용: python check_case.py body.txt (hwpx_dump 결과) |
| `check_more.py` | 바로 사용 | 풍화토·풍화암 판정 기준(50/10)과 다른 SPT 값, 공내재하시험 표의 지층·N값·암종을 4.3.2 공별 표와 대조. 사용: python check_more.py body.txt (hwpx_dump 결과) |
| `check_spt.py` | 바로 사용 | 4.4.1 시추공별 표준관입시험 표 ↔ 4.3.2 공별 지층표의 N값 범위 대조, 횟수·합계 확인. 사용: python check_spt.py body.txt |
| `check_tables.py` | 바로 사용 | 4.3.2 시추조사 결과 자동 대조 — body.txt(hwpx_dump 결과)에서 ① 공별 표(심도·두께 연속성) ② 지층두께 총괄표(표68~70) ③ 지층특성 요약표 ④ 지층개황 서술의 숫자가 서로 맞는지 봄. 사용: python check432.py body.txt |
