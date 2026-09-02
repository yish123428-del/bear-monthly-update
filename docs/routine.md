# Routine 「BEAR 월간 현황 업데이트」 설정

Claude Code Routine 이 매월 실행일 후보일에 이 저장소의 스킬을 돌린다.
Routine 세션에는 **Google Drive 와 Gmail 커넥터가 반드시 켜져 있어야** 한다
(시트 읽기 → `mcp__Google_Drive__read_file_content`, 알림 → `mcp__Gmail__send_message`).

## 현재 구성 (2026-09-02)

- Routine `trig_016Vh1FfExscSXaLcSuZnDQx` — **커넥터를 가진 기존 세션(`session_01MZxPsy8eSauhyxsn6nDsCD`)으로 되돌아오는(self-bind) 방식**.
  세션 안의 `create_trigger` 로 만든 "새 세션" Routine 에는 이 조직에서 커넥터가 붙지 않아(2026-09-02 테스트 발화로 확인: Drive·Gmail 도구 없음)
  스킬을 만든 세션 자체를 매월 재개하도록 바꿨다.
- 그 세션이 사라지거나 컨텍스트가 너무 커지면, claude.ai 의 **Routines 화면에서 아래 값으로 새로 만든다**
  (UI 에서는 커넥터를 직접 고를 수 있다).

| 항목 | 값 |
|---|---|
| 이름 | BEAR 월간 현황 업데이트 |
| 환경 | ander (`env_014tb2wDeRkFwoGJ77EJvbSh`) |
| 스케줄 (UTC cron) | `0 0 28-31,1-10 * *` — 매월 28~31일·1~10일 09:00 KST. 실행일이 아닌 날은 게이트에서 바로 끝난다 |
| 커넥터 | Google Drive, Gmail |
| 알림 | 푸시 |
| 저장소 | `yish123428-del/bear-monthly-update` 기본 브랜치 |

## 프롬프트

```
BEAR 아이디어 제안 월간 현황 업데이트를 실행합니다 (저장소 yish123428-del/bear-monthly-update 의 `bear-monthly-update` 스킬).

절차:
1. 저장소를 준비합니다. 세션에 이미 체크아웃돼 있으면 그대로 쓰고, 없으면 add_repo(owner "yish123428-del", repo "bear-monthly-update", access "push")를 호출한 뒤 안내대로 클론합니다. 기본 브랜치를 사용합니다.
2. 저장소 루트에서 `bash .claude/skills/bear-monthly-update/scripts/bootstrap.sh` 를 실행합니다 (pip 의존성·폰트).
3. `.claude/skills/bear-monthly-update/SKILL.md` 를 읽고 "절차 (Routine · 자동)" 0~8단계를 사람 확인 없이 끝까지 수행합니다.
   - 1단계 게이트 `python3 scripts/run_date.py --check` 가 실행일이 아니라고 하면(exit 3) "오늘은 실행일이 아님, 다음 실행일 …" 한 줄만 보고하고 즉시 종료합니다. 커밋·업로드·메일 없음.
   - 실행일이면: Google Drive 커넥터로 시트 수집(mcp__Google_Drive__read_file_content) → pipeline prepare → review.md 확인·제안내용 문구 다듬기 → pipeline build → data/·output/ 커밋 후 체크아웃된 브랜치로 푸시 → Drive 사본 업로드 → Gmail 알림(mcp__Gmail__send_message, yish123428@gmail.com 한 명에게만, .xlsx·.eml 첨부) → 한 문단 요약.

주의:
- 사업팀 공유 메일(.eml)은 절대 발송하지 않습니다. 파일로만 만듭니다. recipients.json 의 주소로 메일을 보내지 마세요. Google Sheet 에는 쓰지 않습니다.
- 팀 미매칭·누계 불일치·캐치업 행처럼 확신이 없는 항목은 멈추지 말고 산출물을 만들되 review.md 와 알림 제목의 [확인 필요] 로 남깁니다. 팀·제안자·건수 같은 사실은 추측으로 바꾸지 않습니다.
- 베이스 현황판이 data/ 에 없으면(build exit 5) 산출물을 만들지 말고 Gmail 로 "data/ 에 최신 현황판 xlsx 업로드 필요" 를 보내고 종료합니다.
- 시트·스킬·저장소 접근이 실패하면 무엇이 왜 실패했는지 정확히 보고합니다. 추측으로 채우지 마세요.
```

## 실행일 확인

```bash
python3 .claude/skills/bear-monthly-update/scripts/run_date.py --next 12
```

## 첫 실행 전 준비

1. 최신 현황판 `BEAR_아이디어_제안_현황판_YY년M월_vN.xlsx` 를 `data/` 에 커밋한다. 없으면 build 가 exit 5 로 멈춘다.
2. Drive 사본 폴더는 첫 실행이 만들어 `config.json` 의 `drive_upload.folder_id` 에 기록한다.
3. 수동으로 한 번 확인하려면 Routine 을 즉시 발화(fire)하거나, 대화에서 `/bear-monthly-update` 를 실행한다.
