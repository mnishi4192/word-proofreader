# 文書校正アシスタント - manifest.xml 生成スクリプト（Windows 用）

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  文書校正アシスタント セットアップ" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "GitHub Pages の URL を入力してください。"
Write-Host "例: https://yourusername.github.io/word-proofreader"
Write-Host ""
$PagesUrl = Read-Host "GitHub Pages URL"
$PagesUrl = $PagesUrl.TrimEnd('/')

if ([string]::IsNullOrWhiteSpace($PagesUrl)) {
    Write-Host "エラー: URL が入力されていません。" -ForegroundColor Red
    exit 1
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# manifest.xml を生成
$template = Get-Content "$ScriptDir\manifest_template.xml" -Raw -Encoding UTF8
$manifest = $template -replace 'GITHUB_PAGES_URL', $PagesUrl
$manifest | Out-File "$ScriptDir\manifest.xml" -Encoding UTF8

Write-Host ""
Write-Host "✓ manifest.xml を生成しました。" -ForegroundColor Green

# ネットワーク共有フォルダに配置（Word への登録方法）
$SharedFolder = "$env:USERPROFILE\Documents\WordAddins"
if (-not (Test-Path $SharedFolder)) {
    New-Item -ItemType Directory -Path $SharedFolder | Out-Null
}
Copy-Item "$ScriptDir\manifest.xml" "$SharedFolder\manifest.xml" -Force

Write-Host "✓ manifest.xml を $SharedFolder にコピーしました。" -ForegroundColor Green
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  次の手順で Word に登録してください" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Word を起動"
Write-Host "2. 「ファイル」→「オプション」→「トラスト センター」"
Write-Host "   →「トラスト センターの設定」→「信頼できるアドイン カタログ」"
Write-Host "3. カタログの URL に以下を入力して「カタログの追加」:"
Write-Host "   $SharedFolder" -ForegroundColor Yellow
Write-Host "4. 「メニューに表示する」にチェックを入れて OK"
Write-Host "5. Word を再起動"
Write-Host "6. 「挿入」→「アドイン」→「個人用アドイン」→「共有フォルダー」"
Write-Host "   から「文書校正アシスタント」を選択"
Write-Host ""
