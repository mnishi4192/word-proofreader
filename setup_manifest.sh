#!/bin/bash
# ============================================================
# 文書校正アシスタント - manifest.xml セットアップスクリプト (Mac)
# GitHub Pages の URL を入力すると manifest.xml を自動生成します
# ============================================================

echo "======================================"
echo "  文書校正アシスタント セットアップ"
echo "======================================"
echo ""
echo "GitHub Pages の情報を入力してください。"
echo ""

read -p "GitHub ユーザー名（例: tanaka）: " GITHUB_USER
if [ -z "$GITHUB_USER" ]; then
  echo "エラー: GitHub ユーザー名が入力されていません。"
  exit 1
fi

read -p "リポジトリ名（例: word-proofreader）: " REPO_NAME
if [ -z "$REPO_NAME" ]; then
  echo "エラー: リポジトリ名が入力されていません。"
  exit 1
fi

BASE_URL="https://${GITHUB_USER}.github.io/${REPO_NAME}"
echo ""
echo "設定する URL: ${BASE_URL}"
echo ""

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
sed \
  -e "s|YOUR_GITHUB_USERNAME|${GITHUB_USER}|g" \
  -e "s|YOUR_REPO_NAME|${REPO_NAME}|g" \
  "${SCRIPT_DIR}/manifest_template.xml" > "${SCRIPT_DIR}/manifest.xml"

echo "✓ manifest.xml を生成しました。"
echo ""

WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
if [ -d "$HOME/Library/Containers/com.microsoft.Word" ]; then
  mkdir -p "$WEF_DIR"
  cp "${SCRIPT_DIR}/manifest.xml" "$WEF_DIR/manifest.xml"
  echo "✓ manifest.xml を Word に登録しました。"
  echo "  場所: ${WEF_DIR}/manifest.xml"
  echo ""
  echo "Word を再起動すると「ホーム」タブに「文書を校正する」ボタンが表示されます。"
else
  echo "Word が見つかりませんでした。"
  echo "manifest.xml を手動で以下のフォルダにコピーしてください:"
  echo "  ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/"
fi

echo ""
echo "======================================"
echo "  セットアップ完了！"
echo "======================================"
