#!/usr/bin/env bash
# セットアップ（macOS / Linux）: .venv を作り、このツールをインストールする
#   ./setup.sh           利用者
#   ./setup.sh --dev     開発者（テスト・lint も入れる）
set -euo pipefail
REPO="$(cd "$(dirname "$0")" && pwd)"
cd "$REPO"

PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ] || ! "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'; then
  echo "Python 3.10 以上が必要です: https://www.python.org/downloads/" >&2
  exit 1
fi

[ -x .venv/bin/python ] || { echo "== 仮想環境 .venv を作成"; "$PY" -m venv .venv; }
VPY="$REPO/.venv/bin/python"
"$VPY" -m pip install --upgrade pip --quiet || echo "（pip の更新はスキップ）"

echo "== doc-to-markdown をインストール"
"$VPY" scripts/install_deps.py "${1:-}"

.venv/bin/doc2md --check-env
echo
echo "完了しました。"
echo "  - VS Code でこのフォルダを開き、Copilot Chat で / を入力するとプロンプトが使えます"
echo "  - ターミナルで使う場合: source .venv/bin/activate してから doc2md --help"
