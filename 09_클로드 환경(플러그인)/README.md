# 09_클로드 환경(플러그인)

Claude Code에 깔아 쓰는 외부 도구 모음(2026.11 설치). 설치 결과물은 넣지 않고 **무엇을, 어떻게 깔았는지**만 적었다.
새 PC에서는 `install.bat`을 한 번 실행하면 된다.

## 깔아 둔 도구

| 이름 | 종류 | 하는 일 | 켜는 법 |
|---|---|---|---|
| graphify | 스킬 + 파이썬(`graphifyy`) | 폴더 속 코드·문서·그림을 지식 그래프로 만듦. Claude가 파일을 하나씩 뒤지지 않고 그래프부터 봄 | 과업 폴더마다 처음 한 번 `/graphify .` |
| ponytail | 플러그인 | 코드를 꼭 필요한 만큼만 짜게 함(있는 코드·표준 라이브러리 먼저). 오류 처리·검증은 남김 | 자동(항상 켜짐) |
| caveman | 플러그인 | Claude의 **말**(설명 문장)을 짧게 줄여 토큰 절약. 코드·경로·오류문은 그대로 | 세션마다 `/caveman` (lite·full·ultra) |
| hwp-mcp | MCP 서버(jkf87/hwp-mcp, 이 폴더의 `hwp-mcp\`) | 한글 프로그램을 COM으로 조작해 hwp 문서 만들기·글자 넣기·표 만들고 채우기·일괄 작업 | 등록하면 Claude가 필요할 때 씀 |
| commit-commands | 플러그인(Anthropic 공식) | 바뀐 내용을 보고 커밋 메시지를 만들어 커밋까지 해 줌 | `/commit`, `/commit-push-pr` 등 |
| claude-md-management | 플러그인(Anthropic 공식) | CLAUDE.md를 점검하고, 세션에서 배운 요령을 CLAUDE.md에 정리해 넣음 | 필요할 때 명령으로 |
| ralph-loop | 플러그인(Anthropic 공식) | Claude가 멈추려 할 때 같은 지시를 다시 넣어, 완료 조건이 맞을 때까지 반복시킴 | 필요할 때 `/ralph-loop "할 일" ...` |
| headroom | 플러그인(디렉터리, Stefan Petrovici) | 큰 로그·긴 문서·큰 JSON·CSV를 통째로 읽지 않고 개요부터 본 뒤 필요한 부분만 읽음. 줄 번호 붙은 요약을 남김 | 큰 파일을 다룰 때 자동, 직접은 `/peek 파일` |
| claude-code-setup | 플러그인(Anthropic 공식) | 작업 폴더를 분석해 맞는 훅·스킬·MCP·서브에이전트를 추천 | "이 폴더에 맞는 Claude Code 자동화 추천해 줘" |
| frontend-design | 플러그인(Anthropic 공식) | 화면(UI)·웹페이지를 만들 때 뻔한 기본값 대신 의도 있는 디자인(글꼴·색·배치)을 고르게 함 | UI 만들 때 자동 |
| notion | 플러그인(Notion Labs) + Notion MCP | Notion 페이지 만들기·DB 행 추가·검색, 대화 내용을 Notion 문서로 정리 | `/create-page`, `/search` 등. 처음 한 번 Notion 로그인(연결 승인) |

- 줄이는 대상이 서로 다르다: graphify는 파일 찾는 시간, ponytail은 코드 양, caveman은 말 양. ralph-loop는 줄이는 게 아니라 **끝날 때까지 반복**시키는 도구.
- caveman은 보고서 문장을 다듬거나 설명을 자세히 들어야 할 때는 끈다.
- 셋 다 **cmd에서 `claude`로 띄운 Claude Code**에서만 동작한다(웹·일반 채팅에서는 안 됨).
- headroom, claude-code-setup, frontend-design, notion 넷은 2026.11.06 **claude.ai 계정(앱 대화창)에 켜 둔** 플러그인이다. 앱에서는 계정에 붙어 있어 따로 깔 게 없고, cmd의 Claude Code에서도 쓰려면 아래 명령(install.bat)으로 한 번 더 깐다.

## 준비

- Node.js LTS (ponytail 훅이 node를 씀, PATH에 있어야 함)
- Python 3.10 이상 (지금 PC는 3.13)
- Git for Windows (플러그인 마켓플레이스를 GitHub에서 받음)
- Claude Code CLI: `npm install -g @anthropic-ai/claude-code`

## 설치 명령 (install.bat과 같은 내용)

```bat
pip install graphifyy
graphify install

claude plugin marketplace add DietrichGebert/ponytail
claude plugin install ponytail@ponytail

claude plugin marketplace add JuliusBrussee/caveman
claude plugin install caveman@caveman

claude plugin install ralph-loop@anthropic-plugin-directory
claude plugin install commit-commands@anthropic-plugin-directory
claude plugin install claude-md-management@anthropic-plugin-directory
claude plugin install claude-code-setup@anthropic-plugin-directory
claude plugin install frontend-design@anthropic-plugin-directory
claude plugin install headroom@anthropic-plugin-directory
claude plugin install notion@anthropic-plugin-directory

pip install xlwings ezdxf pyproj
```

- Anthropic 공식 플러그인(ralph-loop, commit-commands, claude-md-management, claude-code-setup, frontend-design)이 "not found"이면 `@anthropic-plugin-directory` 대신 `@claude-plugins-official`로. 그래도 안 되면 공식 저장소를 직접 추가: `claude plugin marketplace add anthropics/claude-plugins-official` → `claude plugin install ralph-loop@claude-plugins-official`

- notion이 "not found"이면 Notion 저장소를 직접 추가: `claude plugin marketplace add makenotion/claude-code-notion-plugin` → `claude plugin install notion`. 설치 뒤 Claude Code 안에서 `/mcp` → notion → 로그인.
- headroom은 디렉터리에만 있다. "not found"이면 claude.ai 앱에서 쓰고, CLI에서는 건너뛴다.

- pip 패키지 이름은 **graphifyy**(y 두 개). 명령어·스킬 이름은 graphify.
- `graphify`가 명령으로 안 잡히면 `python -m graphify install`.
- 확인: `claude plugin list` (ponytail·caveman이 enabled), graphify는 플러그인이 아니라 스킬이라 목록에 안 나오는 게 정상 → `%USERPROFILE%\.claude\skills\graphify\` 확인.

## 같이 깐 파이썬 라이브러리

| 이름 | 쓰는 곳 |
|---|---|
| xlwings | 엑셀 COM 작업을 짧고 안전하게 (DispatchEx·숨은 엑셀 정리 등을 대신 처리) |
| ezdxf | 단면도 dxf 만들기·읽기·고치기 (CADian에서 바로 열림) |
| pyproj | 좌표계 변환 (GRS80 TM, 경위도 등) |

## commit-commands로 커밋하기

Claude Code에서 `/commit` → 바뀐 파일을 보고 메시지를 만들어 커밋한다. 푸시는 GitHub Desktop의 **Push origin**이나 `git push`로 한다(이 저장소는 푸시를 사용자가 한다).

## hwp-mcp (한글 조작)

- 한글 프로그램을 **COM**으로 조작한다(마우스를 움직이는 게 아니라 프로그램에 직접 명령). 그래서 이 PC처럼 **한글이 깔린 Windows**에서만 된다.
- 받은 위치: 이 폴더의 `hwp-mcp\` (깃에는 올리지 않음, install.bat이 다시 받음).
- 설치 (cmd, 이 폴더에서):

```bat
pip install -r hwp-mcp\requirements.txt
pip install --force-reinstall pywin32
claude mcp add hwp -s user -- python "D:\SHERPA\10_Toolkit\09_클로드 작업도구\09_클로드 환경(플러그인)\hwp-mcp\hwp_mcp_stdio_server.py"
claude mcp list
```

- `claude mcp list`에서 `hwp`가 ✓면 성공. `-s user`라서 어느 폴더에서 열어도 쓸 수 있다.
- 첫 시험은 **새 문서·사본으로**: "hwp 도구로 새 문서 만들고 '시험'이라고 써서 D:\임시\시험.hwp로 저장해줘".
- "접근 허용" 창이 계속 뜨면 한글 보안 모듈 문제 → `hwp-mcp\security_module\` 안내와 `01_한글/hwp_com.py` 설명 참고.
- COM 오류가 나면 `python -m win32com.client.makepy "HWPFrame.HwpObject"`.
- 보고서 본문은 늘 규칙대로 **사용자가 고쳐 달라고 할 때만** 고친다.

## ralph-loop 쓰는 법

Claude Code 안에서

```
/ralph-loop "엑셀 검토 스크립트 오류 없이 돌아가게 고쳐" --max-iterations 20 --completion-promise "DONE"
```

- `--max-iterations`: 최대 반복 횟수. **꼭 넣는다**(처음엔 10~20).
- `--completion-promise`: Claude가 이 말을 하면 멈춤.
- 멈추기: `/cancel-ralph`
- 계속 도는 구조라 **사용량이 빨리 닳는다.** "오류 없이 실행되면 끝"처럼 완료 기준이 분명한 일에만 쓴다.
- 과업 원본 폴더가 아니라 **사본 폴더**에서 돌린다.

## graphify를 과업 폴더에 붙이기

과업 폴더에서

1. Claude Code 안에서 `/graphify .` → `graphify-out\`에 그래프(html·md·json) 생성
2. cmd에서 `graphify claude install` → 그 폴더의 `CLAUDE.md`에 "그래프 먼저 보기" 규칙과 훅이 들어감

`graphify-out\`은 다시 만들 수 있으니 깃에 올리지 않는다(.gitignore에 넣어 둠).

## 검토하고 안 깐 것 (2026.11.06)

| 이름 | 무엇 | 안 깐 이유 |
|---|---|---|
| OmniRoute | Claude Code 앞에 두는 자체 호스팅 AI 게이트웨이. `ANTHROPIC_BASE_URL`을 돌려 다른 회사 모델로 우회 | 실제 답은 Claude가 아닌 경우가 많고, API 키·대화가 제3자 프록시를 거침. "무료 무제한" 홍보 위주 |
| claude-mem | 세션마다 한 일을 기록·압축해 다음 세션에 넣어 주는 기억 플러그인(thedotmack/claude-mem) | claude.ai 메모리·CLAUDE.md와 겹침. Node 18+·Bun·로컬 워커(37777 포트) 상주 필요, DB가 커짐. 쓰게 되면 공식 저장소로만(코인 홍보 포크 주의) |
| Task Observer | 사용자가 고쳐 준 내용을 CLAUDE.md 규칙으로 쌓음 | claude-md-management·claude.ai 메모리와 겹침. CLAUDE.md에 호출 줄이 없으면 잘 안 붙음. 필요하면 나중에 |

## 어디에 깔리나 (모두 C드라이브)

| 항목 | 위치 |
|---|---|
| 플러그인 | `%USERPROFILE%\.claude\plugins\` |
| 스킬(graphify) | `%USERPROFILE%\.claude\skills\graphify\` |
| 플러그인 켜짐 설정 | `%USERPROFILE%\.claude\settings.json` |
| Claude Code 본체 | `%APPDATA%\npm\` |
| graphify 파이썬 | `...\Python313\Lib\site-packages\` |

**깃에 올리지 않는 것**: `%USERPROFILE%\.claude.json`(계정), `.claude\.credentials.json`(로그인 토큰), `.claude\plugins\` 통째.

## 겪은 일

- `pip install grahpifyy` → 오타라 "No matching distribution". 이름 확인.
- `Ignoring invalid distribution ~ywin32` 경고는 예전에 깔다 만 찌꺼기. 설치에는 영향 없음. 거슬리면 site-packages의 `~`로 시작하는 폴더를 지운다.
- 이 대화창(클라우드 세션)에서 `/plugin`을 치면 "isn't available in this environment" → cmd에서 `claude plugin ...`으로 깐다.
- `claude plugin marketplace add`만 하고 install을 안 하면 `plugin list`가 비어 있다. 두 단계 모두 해야 함.
- caveman 설치 때 "yarn or pnpm lockfile ... not installed" 안내가 뜨지만 기본 기능(짧게 말하기)은 동작한다.
