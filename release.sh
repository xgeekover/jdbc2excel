#!/usr/bin/env bash
# GitHub 릴리스 자동화: 사전 점검 → 테스트 → zip 빌드 → 태그 → 릴리스 생성
#
# 사용법:
#   ./release.sh 1.1.0            # v1.1.0 릴리스
#   ./release.sh 1.1.0 --dry-run  # 태그·푸시·릴리스 없이 점검과 zip 빌드까지만
#
# 릴리스 생성 단계에서 실패해도 태그가 HEAD를 가리키면 그대로 재실행하면 이어서 진행된다.
set -euo pipefail

err() { echo "ERROR: $*" >&2; exit 1; }

VERSION="${1:-}"
[[ "$VERSION" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] \
    || { echo "사용법: ./release.sh <버전> [--dry-run]   (예: ./release.sh 1.1.0)"; exit 1; }
TAG="v$VERSION"
DRY_RUN=false
[[ "${2:-}" == "--dry-run" ]] && DRY_RUN=true

cd "$(dirname "$0")"

echo "=== 1/5 사전 점검 ==="
git diff --quiet && git diff --cached --quiet \
    || err "커밋되지 않은 변경이 있습니다. 커밋 후 다시 실행하세요."
UNTRACKED=$(git ls-files --others --exclude-standard)
[[ -z "$UNTRACKED" ]] \
    || err "추적되지 않은 파일이 있습니다 (git archive가 zip에 담지 않음): ${UNTRACKED//$'\n'/, }"
BRANCH=$(git rev-parse --abbrev-ref HEAD)
[[ "$BRANCH" == "main" ]] || err "main 브랜치가 아닙니다 (현재: $BRANCH)"
git fetch origin main --tags --quiet
[[ "$(git rev-parse HEAD)" == "$(git rev-parse origin/main)" ]] \
    || err "origin/main과 동기화되어 있지 않습니다. push 또는 pull 후 다시 실행하세요."
gh auth status >/dev/null 2>&1 || err "gh 인증이 필요합니다 (gh auth login)"
gh release view "$TAG" >/dev/null 2>&1 && err "릴리스 $TAG가 이미 존재합니다"
if git rev-parse -q --verify "refs/tags/$TAG" >/dev/null; then
    [[ "$(git rev-parse "$TAG^{commit}")" == "$(git rev-parse HEAD)" ]] \
        || err "태그 $TAG가 이미 다른 커밋을 가리킵니다"
    echo "기존 태그 $TAG가 HEAD와 일치 — 재시도로 간주하고 계속 진행"
fi
echo "통과 (브랜치 main, 워킹트리 클린, origin 동기화)"

echo "=== 2/5 테스트 ==="
PY=".venv/bin/python"
[[ -x "$PY" ]] || PY="python3"
"$PY" tests/run_tests.py || err "테스트 실패 — 릴리스를 중단합니다"

echo "=== 3/5 소스 zip 빌드 ==="
WORK_DIR=$(mktemp -d)
trap 'rm -rf "$WORK_DIR"' EXIT
ZIP="$WORK_DIR/jdbc2excel-$VERSION.zip"
git archive --format=zip --prefix="jdbc2excel-$VERSION/" -o "$ZIP" HEAD
echo "생성: $(du -h "$ZIP" | cut -f1) $(basename "$ZIP") ($(unzip -l "$ZIP" | tail -1 | awk '{print $2}')개 파일)"

if $DRY_RUN; then
    echo "=== dry-run 종료 (태그·푸시·릴리스 생성 안 함) ==="
    exit 0
fi

echo "=== 4/5 태그 푸시 ==="
git rev-parse -q --verify "refs/tags/$TAG" >/dev/null || git tag -a "$TAG" -m "$TAG"
git push origin "$TAG"

echo "=== 5/5 릴리스 생성 ==="
gh release create "$TAG" \
    "$ZIP#jdbc2excel-$VERSION.zip (소스 아카이브)" \
    --title "$TAG" \
    --generate-notes

ASSETS=$(gh release view "$TAG" --json assets --jq '[.assets[].name] | join(", ")')
[[ "$ASSETS" == *"jdbc2excel-$VERSION.zip"* ]] || err "릴리스는 생성됐지만 zip 첨부 확인 실패"
echo "완료: $(gh release view "$TAG" --json url --jq .url)"
echo "첨부 자산: $ASSETS"
