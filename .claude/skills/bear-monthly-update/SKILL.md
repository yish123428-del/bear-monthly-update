---
name: bear-monthly-update
description: 신제품기획팀 'BEAR 아이디어 제안' 월간 현황 업데이트 자동화. 매월 첫 월요일(1일이 화요일이면 전월 말일, 공휴일이면 다음 영업일)에 Google Sheet 「Bear 아이디어 제안」을 Drive 커넥터로 직접 읽어 (1) 현황판 엑셀(BEAR_아이디어_제안_현황판_YY년M월_vN.xlsx)을 갱신하고 (2) 사업팀별 공유 이메일(.eml)을 생성해 저장소 output/ 에 커밋하고 본인 Gmail 로 알린다. Routine(자동)과 대화(수동) 양쪽에서 쓴다. 사용자가 "BEAR 아이디어 제안 업데이트", "BEAR 월간 현황 공유", "신규 제안 현황판 업데이트", "BEAR 이메일 만들어줘", "현황판 N월 업데이트", "BEAR 제안 메일 작성", "BEAR 누적 제안 요약 갱신", "/bear-monthly-update" 등 BEAR 월간 보고/공유를 요청할 때 반드시 이 스킬을 사용할 것.
---

# BEAR 아이디어 제안 월간 업데이트

신제품기획팀(이성희)이 매월 1주차에 하는 'BEAR 아이디어 제안' 현황 공유를 자동화한다. 산출물은 둘:

1. **현황판 엑셀** `BEAR_아이디어_제안_현황판_YY년M월_vN.xlsx` — 이전 달 현황판에 신규 제안을 반영
2. **공유 메일 파일** `BEAR_아이디어_제안_현황_공유_YY년M월_1주차.eml` — 기존 메일 서식·서명·하이퍼링크 재현, 시트 캡처 이미지 인라인, 현황판 첨부

## 원칙

- **사업팀 공유 메일은 발송하지 않는다.** `.eml` 은 파일로만 만든다. From/To/Cc 는 `assets/recipients.json` 대로 채우지만 어떤 메일 서버로도 보내지 않는다. 실제 발송은 사용자가 Naver Works 에서 .eml 을 열어 직접 한다.
- Gmail 커넥터로 보내는 것은 **본인(`config.notify.gmail_to`) 앞 완료 알림 1통**뿐이다.
- **원본은 Drive 시트 하나, 베이스는 `data/` 최신 현황판.** 시트에는 쓰지 않는다. 사용자가 올린 파일보다 `data/` 가 우선이다(수동 모드에서 사용자가 파일을 주면 그것을 `--base` 로 쓴다).
- **수집은 Claude, 가공은 파이썬.** MCP 는 파이썬에서 못 부르므로 시트 읽기·Drive 업로드·Gmail 은 Claude 가 하고, 파싱·날짜·엑셀·메일 생성은 스크립트가 한다.
- Routine 에는 사람이 없다. 확신이 없으면 **멈추지 말고 산출물을 만들되** review.md 와 알림 제목에 `[확인 필요]` 를 남긴다. 사실(팀·건수·제안자)은 추측으로 바꾸지 않는다.

## 경로

스킬 루트: `.claude/skills/bear-monthly-update/` (아래 명령은 여기서 실행). 저장소 루트의 `data/`(베이스 현황판, `runs.json`), `output/<YY년M월>/`(산출물), `work/`(중간 파일, 미커밋).

## 절차 (Routine · 자동)

### 0. 부트스트랩

```bash
bash scripts/bootstrap.sh          # pip 의존성(openpyxl, pymupdf, holidays, pillow) + 나눔고딕 → ~/.fonts
python3 -c "import openpyxl, pymupdf, holidays, PIL"
```

실패하면 `pip install -r ../../../requirements.txt` 를 직접 실행한다. SessionStart hook 이 이미 돌렸을 수 있다.

### 1. 실행일 게이트

```bash
python3 scripts/run_date.py --check      # exit 0 = 실행, exit 3 = 오늘은 아님
```

JSON 이 나온다. `run: false` 면 **"오늘(…)은 실행일이 아닙니다. 다음 실행일 … (보고월 …)" 한 줄만 남기고 즉시 종료**한다. 커밋·업로드·메일 없음.
`run: true` 면 `report_month`(예 `26년 9월`), `prev_month`(`8월`), `note`(정규/캐치업)를 기억한다. 규칙은 `references/date-rule.md`.

### 2. 시트 수집

`config.json` 의 `sheet.file_id` 로 `mcp__Google_Drive__read_file_content` 를 호출한다. 결과가 커서 하네스가 파일로 저장하고 경로를 알려준다 → 그 경로가 `--dump`. 드물게 인라인이면 `Write` 로 `work/sheet_dump.json` 에 저장한다.
`mcp__Google_Drive__get_file_metadata` 로 `modifiedTime` 도 받아 둔다.
`download_file_content` 는 쓰지 말 것(첫 탭 CSV 만 나온다). 구조는 `references/sheet-schema.md`.

### 3. prepare — 신규 제안 초안

```bash
python3 scripts/pipeline.py prepare --dump <덤프경로> --report-month "26년 9월" \
  --sheet-modified-time "<modifiedTime>"
```

`work/26년9월/` 에 `new_proposals.xlsx`, `proposals.json`, `review.md` 가 생긴다. 베이스는 `data/` 에서 자동 선택(보고월 이전 최신). **review.md 를 읽는다.**

- `proposals.json` 의 각 항목 중 **`제안내용` 과 `이메일표기` 문구만** 다듬어도 된다. `_source` 의 원문(제품명·성분·소분류·적응증)만으로 한 줄 요약을 쓴다. 기존 문체 예: `우루사+실리스칸 복합제(UDCA+실리마린)`, `이지에프 연고/외용액(rhEGF) 신규 적응증(쇼그렌, 외안성 질환 등)`. 팀·제안자·건수·검토결과·월은 바꾸지 않는다.
- `_unmatched: true`(팀 미매칭)는 그대로 둔다. 시트1 카운트에는 안 들어가고 시트2 행은 들어간다. 알림에 `[확인 필요]` 가 붙는다. 자주 나오는 표기는 나중에 `config.mapping.team_aliases` 에 추가한다.
- 캐치업 행(`_reason` 에 "캐치업")은 이전 달에 누락된 것이므로 포함하되 review 에 남긴다.
- 헤더를 못 찾아 스크립트가 실패하거나 신규 0건이면 build 를 건너뛰고 7단계(알림)로 가서 상황을 보고한다.

### 4. build — 현황판·이미지·메일

```bash
python3 scripts/pipeline.py build --proposals work/26년9월/proposals.json
```

`output/26년9월/` 에 현황판 `vN.xlsx`, `.eml`, `images/`, `proposals.json`, `review.md`, `manifest.json` 이 생기고, 현황판이 `data/` 에 복사되며 `data/runs.json` 에 기록된다. stdout 의 manifest 에 `notify_subject`, `warnings`, `unmatched_teams` 가 있다.
exit 4 = 이미 처리된 달(`--force` 로 vN 증가), exit 5 = 베이스 없음(사용자에게 현황판 업로드 요청).

이미지는 Chromium(HTML) 백엔드로 그린다. `images/status_table.png` 를 한 번 열어 깨진 곳이 없는지 본다.

### 5. git

```bash
cd <저장소 루트>
git add data output .claude/skills/bear-monthly-update/config.json
git commit -m "BEAR 26년 9월 현황 업데이트 (신규 N건, 총 M건)"
git push origin HEAD || git push -u origin HEAD:bear/26년9월
```

체크아웃된 브랜치(저장소 기본 브랜치)에 그대로 푸시한다. 거부되면 `bear/<월>` 브랜치로 올리고 알림에 "병합 필요" 를 적는다.

### 6. Drive 사본

`config.drive_upload.enabled` 가 true 면 현황판 xlsx 를 `mcp__Google_Drive__create_file` 로 올린다:
`base64Content`(파일 base64), `contentMimeType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, `disableConversionToGoogleType: true`, `parentId: folder_id`, `title: 파일명`.
`folder_id` 가 비어 있으면 먼저 `mimeType: application/vnd.google-apps.folder` 로 `folder_title` 폴더를 만들고 그 ID 를 `config.json` 에 적어 함께 커밋한다. 실패해도 계속 진행하고 알림에 남긴다.

### 7. 본인 알림 (Gmail)

`mcp__Gmail__send_message` 를 **`config.notify.gmail_to` 한 명에게만** 보낸다.

- 제목: manifest 의 `notify_subject` (`[확인 필요] [BEAR 자동화] 26년 9월 현황판·메일 생성 완료 (신규 N건, 총 M건)`)
- 본문: 신규 제안 목록 한 줄씩, review.md 의 "확인 필요" 절 전문, 저장소 커밋/브랜치, Drive 링크, 실행일 note
- 첨부: 현황판 `.xlsx` 와 `.eml` (base64, `mimeType` 각각 xlsx / `message/rfc822`)

recipients.json 의 사업팀·Cc 주소로는 절대 보내지 않는다.

### 8. 마무리

한 문단으로 요약한다(푸시 알림 본문이 된다): 보고월, 신규/총 건수, 확인 필요 항목 수, 파일 위치.

## 수동 모드 (대화에서 요청)

"9월 업데이트", "BEAR 메일 만들어줘" 처럼 요청하면 게이트를 생략한다.

1. 보고월을 확인한다(예 `26년 9월`). 사용자가 현황판이나 신규 제안 파일을 올리면 `--base` / `--dump` 로 쓴다. 안 올리면 시트는 Drive 커넥터로 읽고 베이스는 `data/` 최신본을 쓴다.
2. `prepare` 후 review.md 와 신규 제안 목록, 누계(`이전달 N건 총 M건`), 미매칭 팀, 수신자 변경 여부를 **사용자에게 확인**한다.
3. `build`. 이미 처리된 달이면 `--force`(v2, v3 …).
4. 산출물 경로를 알려 준다. 커밋·Drive·Gmail 은 사용자가 원할 때만.

## 자주 빠지는 함정

- **팀명 표기 차이** — 시트는 `소화기2`, `호피안`, 현황판은 `소화기2사업팀`, `호흡피부안과사업팀`. 스크립트가 접미어와 alias 로 맞추고, 못 맞추면 `_unmatched`. alias 는 `config.mapping.team_aliases`.
- **제안자 = 시트의 접수자.** 시트 `제안자` 열은 KOL 이다.
- **누계** — 시트1 `[ 누계 진행 결과 _N월 X건 총 Y건 ]` 과 메일 `N월 X건 총 Y건 접수` 는 같은 숫자여야 한다. pipeline 이 한 값으로 넣는다. 베이스 Y + 신규 ≠ 시트 행 수면 review 에 경고 — 시트에 늦게 추가·삭제된 행이 있는지 본다.
- **=SUM 수식** — 시트1 `제안 합계`/`누계` 는 수식. 스크립트는 건드리지 않고 Excel 이 재계산한다. 이미지 렌더러는 SUM 을 직접 평가한다.
- **연도 롤오버** — 시트1 `26년` 월 컬럼과 시트2 `(26.01~)` 는 27년 2월 보고부터 새 레이아웃이 필요하다(사용자 템플릿 작업).
- **전월 말일 실행** — 1일이 화요일인 달은 전월 말일에 돌아 그날 오후 입력분을 놓칠 수 있다. 다음 달 캐치업이 흡수한다.
- **LibreOffice** — 샌드박스에는 Calc 가 없어 `soffice` 로 시트를 못 연다. 렌더러는 Chromium 을 쓴다(`--backend soffice` 는 Calc 가 있을 때만).

## 파일

- `config.json` — 시트 ID, 매핑, 휴일, 알림, 경로
- `scripts/run_date.py` 실행일 · `sheet_to_proposals.py` 덤프→신규 제안 · `pipeline.py` prepare/build · `update_status_excel.py` 현황판 갱신 · `render_sheet_images.py` + `xlsx_html.py` 이미지 · `build_email_eml.py` .eml · `naming.py` 파일명 · `dumpio.py` 덤프 파싱 · `bootstrap.sh`
- `references/sheet-schema.md` 시트 구조 · `date-rule.md` 실행일 · `input_format.md` 신규 제안 입력 사양
- `assets/recipients.json` 수신자·서명 · `email_template.html` 본문 · `fonts/` 나눔고딕
