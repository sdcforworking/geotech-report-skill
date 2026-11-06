# 08_수량산출

예시 수량산출서 양식(시추 수량산출·시험 수량산출 수행 시트)을 따라 과업 수량산출서를 만든다. 자료는 본문 hwpx를 푼 body.txt(지층두께 총괄표·SPT·공내재하·토질시험 표)와 `03_주상도/xsm_size.py`의 규격 json.

```
python ../01_한글(hwp·hwpx)/hwpx_dump.py 본문.hwpx body.txt
python ../03_주상도(xsm)/xsm_size.py 통합.xsm size.json
python make_qty.py body.txt size.json 수량산출서.xlsx
```

도급 수량은 스크립트 안 표(본문 1.3 값)를 과업에 맞게 고친다. 지층 열: 모래층·자갈층 = 붕적층, 풍화토층 따로.

| 파일 | 종류 | 설명 |
|---|---|---|
| `make_qty.py` | 바로 사용 | 이번 과업(동해안~신가평 10공구) 수량산출서 만들기 — 예시 '수량산출서.xlsx'의 시추·시험 수량산출(수행) 양식을 따름. 자료: 본문 hwpx(2026.11.02(2))를 풀어 쓴 body.txt(지층두께 총괄표·SPT·공내재하·토질시험 표), 공별 규격은 통합 주상도 xsm에서 뽑은 size.json. 병합 칸 없이 만들고(제목은 '선택 영역 가운데'), 합계는 함수로 둠. 사용: python make_qty.py body.txt size.json 출력.xlsx |
