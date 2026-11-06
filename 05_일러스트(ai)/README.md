# 05_일러스트(ai)

일러스트레이터 없이 ai 안 글자를 보고 바꾼다. 자세한 원리는 `ai_text.py` 설명.

```
python ai_text.py list 파일.ai 벽학산
python ai_text.py replace 원본사본.ai 새파일.ai 434 "벽학=오음" "331.1=929.4" --font C:\Windows\Fonts\malgunbd.ttf
```

`ai_replace_example.py`는 PDF 미리보기까지 바꾼 실제 예(글꼴 조각에 없는 '음' 글리프를 채워 넣고, 흰 테두리를 새 글자 윤곽으로 다시 그림)이다. 표·글상자 번호가 박혀 있다.

| 파일 | 종류 | 설명 |
|---|---|---|
| `ai_replace_example.py` | 과업 예시 | 지형및지질 ai 사본에서 산계·수계도(2번째 대지) 오른쪽 위 '벽학산 (331.1m)' 글상자를 '오음산 (929.4m)'으로 바꿈. ① 일러스트 원본 데이터(AIPrivateData, zstd): 텍스트 문서의 해당 글상자(story 434) 글자와 글리프 번호를 같은 길이로 바꾼 뒤 다시 압축 ② PDF 미리보기: 글꼴 조각에 '음' 글리프를 맑은고딕 굵게에서 채워 넣고, 해당 글자 그리기와 흰 테두리를 새 글자로 바꿈 아래쪽 '벽학산 (331.1m)'(story 433)은 건드리지 않음. 사용: python make_new_ai.py 원본사본.ai 새파일.ai |
| `ai_text.py` | 바로 사용 | 일러스트레이터 없이 .ai(PDF 호환, Illustrator 24 이상 zstd 압축) 안의 글상자 글자를 보거나 바꿈. |
| `explore_td.py` | 과업 예시 | (설명 없음 — 코드 참고) |
| `show_td.py` | 바로 사용 | textdoc.bin(일러스트 텍스트 문서) 특정 위치 앞뒤를 사람이 읽을 수 있게 보여 줌 — (þÿ…) 문자열은 UTF-16으로 풀어서 표시. |
| `tdparse.py` | 바로 사용 | 일러스트 AI11 텍스트 문서(PostScript 비슷한 사전 형식) 읽기 — 값마다 원본 바이트 위치(시작·끝)를 같이 기억함. |
| `zstd_ct.py` | 바로 사용 | Git for Windows에 들어 있는 libzstd.dll을 ctypes로 써서 zstd 압축을 풀고 다시 압축함 (파이썬 패키지 설치 없이). |
