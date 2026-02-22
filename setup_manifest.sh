#!/bin/bash
# 文書校正アシスタント - manifest.xml 生成スクリプト（Mac 用）

echo "========================================"
echo "  文書校正アシスタント セットアップ"
echo "========================================"
echo ""
echo "GitHub Pages の URL を入力してください。"
echo "例: https://yourusername.github.io/word-proofreader"
echo ""
read -p "GitHub Pages URL: " PAGES_URL

# 末尾スラッシュを除去
PAGES_URL="${PAGES_URL%/}"

if [ -z "$PAGES_URL" ]; then
  echo "エラー: URL が入力されていません。"
  exit 1
fi

# manifest.xml を生成
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
sed "s|GITHUB_PAGES_URL|${PAGES_URL}|g" \
    "${SCRIPT_DIR}/manifest_template.xml" \
    > "${SCRIPT_DIR}/manifest.xml"

echo ""
echo "✓ manifest.xml を生成しました。"
echo ""
echo "次に、manifest.xml を Word に登録します..."
echo ""

# Word アドインフォルダにコピー
ADDIN_DIR="${HOME}/Library/Containers/com.microsoft.Word/Data/Documents/wef"
mkdir -p "${ADDIN_DIR}"
cp "${SCRIPT_DIR}/manifest.xml" "${ADDIN_DIR}/manifest.xml"

echo "✓ manifest.xml を Word のアドインフォルダにコピーしました。"
echo "  場所: ${ADDIN_DIR}/manifest.xml"
echo ""
echo "========================================"
echo "  セットアップ完了！"
echo "========================================"
echo ""
echo "Word を起動（または再起動）し、"
echo "「挿入」→「アドイン」→「個人用アドイン」から"
echo "「文書校正アシスタント」を選択してください。"
echo ""
