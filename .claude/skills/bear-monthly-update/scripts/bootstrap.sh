#!/usr/bin/env bash
# 세션 부트스트랩: 파이썬 의존성 설치 + 한글 폰트 설치.
# SessionStart hook 과 SKILL.md 0단계에서 호출한다. 실패해도 세션을 막지 않도록 항상 exit 0.
set -u
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_DIR="$(cd "$HERE/.." && pwd)"
REPO_ROOT="$(cd "$SKILL_DIR/../../.." && pwd)"

echo "[bootstrap] python deps"
if ! python3 -c "import openpyxl, pymupdf, holidays, PIL" >/dev/null 2>&1; then
  pip install -q -r "$REPO_ROOT/requirements.txt" 2>&1 | grep -v "WARNING: Running pip as the 'root'" || true
fi
python3 -c "import openpyxl, pymupdf, holidays, PIL; print('[bootstrap] deps ok')" 2>&1 || echo "[bootstrap] deps 설치 실패 — SKILL.md 0단계에서 다시 시도"

echo "[bootstrap] fonts"
mkdir -p "$HOME/.fonts"
for f in NanumGothic.ttf NanumGothicBold.ttf; do
  [ -f "$HOME/.fonts/$f" ] || cp "$SKILL_DIR/assets/fonts/$f" "$HOME/.fonts/$f" 2>/dev/null || true
done
command -v fc-cache >/dev/null 2>&1 && fc-cache -f "$HOME/.fonts" >/dev/null 2>&1 || true

# 시트 이미지 렌더링 백엔드 안내 (Chromium 우선, LibreOffice Calc 는 있으면 사용)
if ls /opt/pw-browsers/chromium*/chrome-linux*/chrome >/dev/null 2>&1 || command -v chromium >/dev/null 2>&1; then
  echo "[bootstrap] renderer: chromium 사용 가능"
else
  echo "[bootstrap] renderer: chromium 없음 — soffice(calc) 필요"
fi
exit 0
