#!/bin/bash
# ============================================================
# office.js をローカルに取得するスクリプト（Mac / Linux 用）
# ============================================================
# Microsoft の CDN から office.js をダウンロードし、
# このフォルダに保存します。
# 一度だけ実行すれば OK です（以降はインターネット不要）。
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
OUT="$SCRIPT_DIR/office.js"
URL="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"

echo "========================================"
echo "  office.js ダウンロード"
echo "========================================"
echo ""
echo "取得先: $URL"
echo "保存先: $OUT"
echo ""

if [ -f "$OUT" ]; then
  echo "既に office.js が存在します。上書きしますか？ [y/N]"
  read -r ans
  if [ "$ans" != "y" ] && [ "$ans" != "Y" ]; then
    echo "スキップしました。"
    exit 0
  fi
fi

curl -fsSL "$URL" -o "$OUT"

echo ""
echo "✓ office.js を取得しました（$(wc -c < "$OUT" | tr -d ' ') バイト）"
echo ""
echo "次のステップ:"
echo "  caddy trust          # CA 証明書を OS に登録（初回のみ）"
echo "  caddy run --config Caddyfile"
echo ""
