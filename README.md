# bear-monthly-update — BEAR 아이디어 제안 월간 현황 자동화

신제품기획팀의 'BEAR 아이디어 제안' 월간 현황 공유를 자동화하는 Claude Code 스킬.

매월 첫 월요일(1일이 화요일이면 전월 말일, 공휴일이면 다음 영업일) 09:00 에 Routine 이 깨어나
Google Sheet 「Bear 아이디어 제안」을 읽고 **현황판 엑셀**과 **사업팀별 공유 메일(.eml)** 을 만들어
`output/` 에 커밋하고 본인 Gmail 로 알린다. 공유 메일은 발송하지 않는다 — 사용자가 .eml 을 열어 직접 보낸다.

## 왜 필요한가

이전에는 매달 (1) 시트를 내려받아 채팅에 올리고 (2) 지난달 현황판도 올리고 (3) 날짜를 기억해야 했다.
이제 시트는 Drive 커넥터로 직접 읽고, 지난달 현황판은 저장소 `data/` 에서 가져오며, 실행일은 스크립트가 계산한다.

## 사용법

```
/bear-monthly-update            # 대화에서 수동 실행 (보고월 확인 → prepare → 검토 → build)
python3 .claude/skills/bear-monthly-update/scripts/run_date.py --next 12   # 향후 실행일
```

자동 실행은 Claude Code Routine `BEAR 월간 현황 업데이트` 가 한다 (cron `0 0 28-31,1-10 * *` UTC).
후보일마다 깨어나 `run_date.py --check` 가 실행일이라고 할 때만 전체를 돈다.

## 구조

```
.claude/skills/bear-monthly-update/
├── SKILL.md                    Claude 오케스트레이션 지침 (Routine · 수동 모드)
├── config.json                 시트 ID · 팀명 매핑 · 휴일 · 알림 · 경로
├── references/
│   ├── sheet-schema.md         Google Sheet 컬럼과 변환 규칙
│   ├── date-rule.md            실행일 규칙 + 2026~2027 표
│   └── input_format.md         신규 제안 입력 xlsx 사양
├── assets/                     메일 템플릿 · 수신자 · 나눔고딕 폰트
└── scripts/
    ├── bootstrap.sh            pip 의존성 + 폰트 설치 (SessionStart hook)
    ├── run_date.py             실행일 게이트 (공휴일·캐치업)
    ├── sheet_to_proposals.py   시트 덤프 → 신규 제안 (new_proposals.xlsx / proposals.json / review.md)
    ├── pipeline.py             prepare / build 오케스트레이터
    ├── update_status_excel.py  현황판 엑셀 갱신
    ├── render_sheet_images.py  시트 → PNG (Chromium HTML 백엔드, LibreOffice 폴백)
    ├── xlsx_html.py            openpyxl → HTML (+ SUM 수식 평가)
    ├── build_email_eml.py      .eml 생성
    ├── naming.py               파일명 · 최신 베이스 · 버전
    └── dumpio.py               Drive 덤프(markdown) 파싱
data/                           베이스 현황판 이력 + runs.json (커밋)
output/<YY년M월>/               월별 산출물 (커밋)
tests/                          unittest + fixtures
```

MCP 도구는 파이썬에서 호출할 수 없다. 그래서 **수집은 Claude, 가공은 파이썬**으로 나눴다.
결정적인 부분(파싱·날짜 계산·엑셀·메일)이 전부 스크립트에 있어 재현·테스트가 된다.

## 처음 한 번

1. 현재 최신 현황판 `BEAR_아이디어_제안_현황판_YY년M월_vN.xlsx` 를 `data/` 에 넣고 커밋한다.
2. `pip install -r requirements.txt` (Routine 세션은 SessionStart hook 이 한다).
3. Drive 사본 폴더는 첫 실행이 만들어 `config.json` 의 `drive_upload.folder_id` 에 기록한다.

## 테스트

```bash
python3 -m unittest discover -s tests
```

날짜 규칙은 stdlib 만으로, 나머지는 `openpyxl` 이 있을 때 돈다. 렌더링(Chromium)은 테스트에서 생략한다.

## 하지 않는 것

사업팀 공유 메일 발송 · Google Sheet 쓰기. 산출물은 파일과 본인 앞 알림 메일뿐이다.
