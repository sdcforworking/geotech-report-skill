# 클로드 작업도구

지반조사 보고서 작업(2026.10~11, 500kV HVDC 동해안~신가평 10공구)에서 Claude와 같이 만든 **파이썬 도구와 작업 요령** 모음.
결과물(보고서·엑셀·그림)은 들어 있지 않고, 다음 과업에서 다시 쓸 **방법**만 모았다.

## 폴더

| 폴더 | 하는 일 | 바로 쓰는 도구 |
|---|---|---|
| `01_한글(hwp·hwpx)` | hwpx를 표·수식까지 글자로 풀기, hwp 표·수식 읽기, 한글 COM 시작 | `hwpx_dump.py`, `hwp_tbl.py`, `hwp_eq.py`, `hwp_com.py` |
| `02_보고서 검토` | 보고서 표끼리 숫자 자동 대조 (공별 표·총괄표·요약표·서술, SPT, 공내재하, 설계사례) | `check_tables.py`, `check_spt.py`, `check_more.py`, `check_case.py` |
| `03_주상도(xsm)` | XsiteMap xsm에서 공번·규격·지층·N값 뽑기, xsm 고치기 | `xsm_size.py`, `xsmparse.py` |
| `04_엑셀` | 엑셀 COM으로 안전하게 고치기, PDF로 내보내기, 이동 내역 기록 | `log_move.py`, `xl_pdf.py` |
| `05_일러스트(ai)` | 일러스트레이터 없이 ai 안 글자 보기·바꾸기 | `ai_text.py` |
| `06_단면도·삽도` | 지층단면도 ai·dxf 만들기, 칸 크기 맞추기, N값-심도 그래프, 한글 칸에 그림 넣기 | (과업 예시) |
| `07_설계지반정수` | 대표 N값, φ·E 경험식, 문헌·사례 범위, 산정결과 시트 만들기 | (과업 예시) |
| `08_수량산출` | 예시 양식으로 수량산출서(총괄·시추·시험) 만들기 | `make_qty.py` |
| `09_클로드 환경(플러그인)` | Claude Code에 깐 Claude Code 플러그인·MCP(graphify, ponytail, caveman, ralph-loop, commit-commands, claude-md-management, claude-code-setup, frontend-design, headroom, notion, hwp-mcp)와 파이썬 라이브러리(xlwings, ezdxf, pyproj) 설치법과 특징, 새 PC용 설치 스크립트 | `install.bat` |
| `_원본` | 세션별 임시 스크립트 전부(디버그용 포함) — 찾아볼 때만. **공개본에는 없음**(과업 자료·개인 정보가 들어 있어 비공개 저장소에만 둠) | |

**바로 쓰는 도구**는 인자로 파일을 받게 다듬은 것이다. **과업 예시**는 그 과업의 시트 이름·표 번호·값이 박혀 있어서, 위쪽 상수를 새 과업에 맞게 바꿔서 쓴다.

## 스킬로 쓰기

이 폴더는 Claude Code 스킬 **geotech-report** 이기도 하다.

- `SKILL.md` — Claude가 읽는 안내서 (항상 지킬 것, 일하는 순서, 도구 목록, 판단 기준)
- `references/` — 길게 참고할 규칙 (예: `xsm편집규칙.md`)
- `user-CLAUDE.md` — 모든 대화 공통 규칙의 원본. 사용자 폴더의 `.claude\CLAUDE.md`로 복사해서 쓴다

새 PC에서 쓰기: 이 저장소를 Clone한 뒤 PowerShell에서 (경로는 Clone한 폴더로)

```powershell
New-Item -ItemType Junction -Path "$env:USERPROFILE\.claude\skills\geotech-report" -Target "D:\SHERPA\10_Toolkit\09_클로드 작업도구"
Copy-Item "D:\SHERPA\10_Toolkit\09_클로드 작업도구\user-CLAUDE.md" "$env:USERPROFILE\.claude\CLAUDE.md"
```

## 준비

- Python 3.13, `pip install pymupdf openpyxl pywin32 pillow`
- Git for Windows (ai 도구가 `libzstd.dll`을 씀)
- 한글 COM을 쓰려면 보안 모듈 등록 (`01_한글/hwp_com.py` 설명 참고)
- 실행은 `python -I 스크립트.py …` (사용자 폴더의 같은 이름 모듈이 끼어들지 않게). 스크립트는 첫머리에서 `sys.stdout.reconfigure(encoding='utf-8')`로 한글 출력을 맞춘다.

## 늘 지키는 작업 규칙

1. **사본으로 작업하고, 덮어쓰기 전에 백업.** 백업은 과업의 `보관/` 아래에 두고, 옮기기 전후 MD5가 같은지 확인한다.
2. **옮기거나 백업하면 이동 내역 엑셀에 한 줄.** `04_엑셀/log_move.py --log 이동내역.xlsx 원래 새 "비고"`
3. **사용자가 파일을 열어 두었는지 먼저 확인.** 같은 폴더의 `~$파일이름` 잠금 파일과 MD5를 보고, 바뀌었으면 사용자 판을 기준으로 다시 한다.
4. **엑셀 출력에는 병합 칸을 쓰지 않는다.** 제목은 '선택 영역의 가운데로'(`centerContinuous`)를 쓴다.
5. **한글 파일은 사용자가 고쳐 달라고 할 때만 고친다.** 보통은 고칠 문장만 만들어 주고 사용자가 붙여 넣는다.

## 요령 모음

### 엑셀
- COM은 `win32com.client.DispatchEx('Excel.Application')`(사용자 엑셀과 분리), `Visible=False`, `DisplayAlerts=False`. 끝나면 `Quit()`. 남은 숨은 엑셀은 명령줄에 `-Embedding`이 있는 것만 종료한다(사용자 엑셀은 건드리지 않음).
- **저장 전 선택된 탭이 하나인지 확인.** openpyxl로 저장하면 탭 여러 개가 선택된 채 남아 '그룹' 상태로 열리고, 사용자가 붙여 넣은 값이 여러 시트에 들어간다. `ws.Activate(); ws.Select()` 후 `xl.ActiveWindow.SelectedSheets.Count == 1` 확인.
- COM으로 열·행을 넣고 빼면 참조가 따라온다. 다른 시트가 참조하는 시트는 지우고 다시 만들지 말고 칸만 고친다(지우면 #REF!).
- `Worksheets.Add(None, 기준시트)`처럼 위치 인자로 넣는다(`After=` 키워드는 맨 앞에 들어감).
- 한글판 엑셀은 NumberFormat `'General'`을 거부한다.
- TEXTJOIN(IF()) 같은 배열 수식은 `Range.FormulaArray`로 넣는다.
- openpyxl로 MINIFS·MAXIFS·AGGREGATE를 쓰려면 `_xlfn.` 접두어가 필요하다(없으면 #NAME?).
- openpyxl로 만든 파일은 계산값이 없으니 한 번 COM으로 열어 `CalculateFull()` 후 저장하면 다른 프로그램에서도 값이 보인다.

### 한글
- hwpx는 zip이다. `Contents/section*.xml`의 `hp:p`(문단)·`hp:tbl`(표, `hp:cellAddr`·`hp:cellSpan`)·`hp:equation/hp:script`(수식)를 읽으면 표 구조와 수식까지 뽑힌다 → `hwpx_dump.py`.
- hwp(5.x)는 OLE 파일. `BodyText/Section*` 스트림을 zlib(-15)로 풀고 레코드를 읽는다(태그 67 글자, 72 표 칸 머리, 88 수식).
- 한글 COM은 `RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')`을 먼저 불러야 '접근 허용' 창이 안 뜬다.

### 일러스트(ai)
- 이 PC에는 일러스트레이터가 없다. ai(PDF 호환)는 PyMuPDF로 열어 그림으로 뽑을 수 있다.
- 글자를 바꾸려면 PDF 쪽이 아니라 **일러스트 원본 자료(AIPrivateData, zstd)** 를 고쳐야 일러스트레이터에서 보인다 → `ai_text.py`. 글자 수가 같을 때만 안전하다.
- 지도 바탕이 대지 경계에서 잘려 있으면(클리핑) 대지 밖으로 늘릴 수 없다. 크기를 바꾸려면 여백을 붙이거나 일러스트레이터에서 넓혀야 한다.

### 주상도(xsm)
- xsm은 바이너리지만 공번·규격·지층 이름은 cp949 글자로 들어 있어 정규식으로 찾을 수 있다.
- 시추 규격(BX/NX)은 공번 바로 뒤 머리 자료에 있다 → `xsm_size.py`. NX 공은 보통 공내재하시험 공이다.

### 지반 자료
- 관입불능 = 50회 타격에 관입량 30cm 미만. 30cm 환산 N = 타격수 × 30 / 관입량.
- 풍화토/풍화암 기준 N = 50/10 (풍화암 판정이 시료 관찰로 50/11인 경우도 있으니 의도 확인).
- 상대밀도: ≤4 매우느슨, 4~10 느슨, 10~30 보통조밀, 30~50 조밀, >50 매우조밀.
- 내부마찰각: 도로교 √(15N)+15(≤45), Dunham √(12N)+C(15~25), Peck 0.3N+27, Ohsaki √(20N)+15.
- 변형계수: Schmertmann αN(실트 0.4·세립~중립모래 0.7·조립모래 1.0·자갈 1.2~1.5), Bowles(모래질 실트 0.30(N+6), 점토질 모래 0.32(N+15), 모래 0.50(N+15), 자갈질 모래 1.20(N+6)), Hisatake 0.5N+7, 도로교 2.8N, 한국지반공학회 1.2N. 평균은 최대값을 빼고 낸다.

### 셸·경로
- 한글·공백·`~`·`#`이 든 경로가 많다. PowerShell에서는 `-LiteralPath`, bash에서는 따옴표.
- PowerShell에서 파이썬 인자에 큰따옴표가 든 문자열을 넘기면 잘린다 → 파일로 넘기거나 파이썬 안에서 처리.
- 콘솔은 cp949라 파이썬 출력이 깨지면 `sys.stdout.reconfigure(encoding='utf-8')`.
