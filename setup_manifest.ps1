# ============================================================
# 文書校正アシスタント - manifest.xml セットアップスクリプト (Windows)
# GitHub Pages の URL を入力すると manifest.xml を自動生成します
# ============================================================

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  文書校正アシスタント セットアップ" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "GitHub Pages の情報を入力してください。"
Write-Host ""

$GitHubUser = Read-Host "GitHub ユーザー名（例: tanaka）"
if ([string]::IsNullOrWhiteSpace($GitHubUser)) {
    Write-Host "エラー: GitHub ユーザー名が入力されていません。" -ForegroundColor Red
    exit 1
}

$RepoName = Read-Host "リポジトリ名（例: word-proofreader）"
if ([string]::IsNullOrWhiteSpace($RepoName)) {
    Write-Host "エラー: リポジトリ名が入力されていません。" -ForegroundColor Red
    exit 1
}

$BaseUrl = "https://$GitHubUser.github.io/$RepoName"
Write-Host ""
Write-Host "設定する URL: $BaseUrl" -ForegroundColor Yellow
Write-Host ""

$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$TemplatePath = Join-Path $ScriptDir "manifest_template.xml"
$OutputPath   = Join-Path $ScriptDir "manifest.xml"

$content = Get-Content $TemplatePath -Raw -Encoding UTF8
$content = $content -replace "YOUR_GITHUB_USERNAME", $GitHubUser
$content = $content -replace "YOUR_REPO_NAME", $RepoName
$content | Set-Content $OutputPath -Encoding UTF8

Write-Host "✓ manifest.xml を生成しました。" -ForegroundColor Green
Write-Host ""

$WefDir = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (-not (Test-Path $WefDir)) {
    New-Item -ItemType Directory -Path $WefDir -Force | Out-Null
}
Copy-Item $OutputPath (Join-Path $WefDir "manifest.xml") -Force
Write-Host "✓ manifest.xml を Word に登録しました。" -ForegroundColor Green
Write-Host "  場所: $WefDir\manifest.xml"
Write-Host ""
Write-Host "Word を再起動すると「ホーム」タブに「文書を校正する」ボタンが表示されます。"
Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  セットアップ完了！" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
