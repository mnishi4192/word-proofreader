#!/bin/bash
# ============================================================
# 文書校正アシスタント（Caddy ローカルホスト版）
# セットアップスクリプト（Mac 用）
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
MANIFEST_SRC="$SCRIPT_DIR/manifest_local.xml"
ADDIN_DIR="${HOME}/Library/Containers/com.microsoft.Word/Data/Documents/wef"

echo "========================================"
echo "  文書校正アシスタント セットアップ"
echo "  （Caddy ローカルホスト版）"
echo "========================================"
echo ""

# ---- Step 1: office.js の確認 ----
if [ ! -f "$SCRIPT_DIR/office.js" ]; then
  echo "⚠️  office.js が見つかりません。"
  echo "今すぐダウンロードしますか？ [y/N]"
  read -r ans
  if [ "$ans" = "y" ] || [ "$ans" = "Y" ]; then
    chmod +x "$SCRIPT_DIR/fetch_office_js.sh"
    "$SCRIPT_DIR/fetch_office_js.sh"
  else
    echo "セットアップを中断しました。office.js を取得してから再実行してください。"
    exit 1
  fi
else
  echo "✓ office.js を確認しました。"
fi

# ---- Step 2: Caddy のインストール確認 ----
if ! command -v caddy &>/dev/null; then
  echo ""
  echo "⚠️  Caddy がインストールされていません。"
  echo "   brew install caddy を実行してからこのスクリプトを再実行してください。"
  exit 1
else
  echo "✓ Caddy を確認しました（$(caddy version 2>&1 | head -1)）"
fi

# ---- Step 3: Caddy をバックグラウンドで起動して CA 証明書を登録 ----
# caddy trust は Caddy の管理 API（port 2019）に接続するため、
# 先に caddy start でバックグラウンド起動が必要。
echo ""
echo "Caddy を一時起動して CA 証明書を OS に登録します..."
echo "管理者パスワードの入力を求められる場合があります。"
echo ""

# 既存の Caddy プロセスがあれば停止
caddy stop 2>/dev/null || true

# Caddyfile のあるフォルダで起動
cd "$SCRIPT_DIR"
caddy start --config Caddyfile
sleep 2  # 起動待ち

# CA 証明書を OS のキーチェーンに登録
sudo caddy trust

echo ""
echo "✓ CA 証明書を登録しました。"
echo ""
echo "⚠️  Mac の場合、キーチェーンアクセスで追加の設定が必要です:"
echo "   1. 「キーチェーンアクセス」を開く"
echo "   2. 「システム」キーチェーンで「Caddy Local Authority」を探す"
echo "   3. ダブルクリック →「信頼」→「常に信頼」に設定"
echo ""
read -p "キーチェーンの設定が完了したら Enter を押してください..."

# ---- Step 4: manifest.xml を Word に登録 ----
mkdir -p "$ADDIN_DIR"
cp "$MANIFEST_SRC" "$ADDIN_DIR/manifest_local.xml"

echo ""
echo "✓ manifest_local.xml を Word のアドインフォルダにコピーしました。"
echo "  場所: $ADDIN_DIR/manifest_local.xml"

# ---- Step 5: Caddy を再起動して通常運用モードへ ----
caddy stop 2>/dev/null || true
echo ""
echo "========================================"
echo "  セットアップ完了！"
echo "========================================"
echo ""
echo "Caddy を起動してください:"
echo "  cd \"$SCRIPT_DIR\""
echo "  caddy run --config Caddyfile"
echo ""
echo "PC 起動時に自動起動する場合:"
echo "  brew services start caddy"
echo "  ※ Caddyfile の root ディレクティブを絶対パスに変更してから実行してください"
echo "  例: root * $SCRIPT_DIR"
echo ""
echo "Caddy 起動後、Word を再起動して"
echo "「挿入」→「アドイン」→「個人用アドイン」から"
echo "「文書校正アシスタント（ローカル）」を選択してください。"
echo ""
