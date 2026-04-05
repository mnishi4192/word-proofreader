# ============================================================
# 文書校正アシスタント（Caddy ローカルホスト版）
# セットアップスクリプト（Windows 用）
# ※ 管理者権限の PowerShell で実行してください
# ============================================================

$ScriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$ManifestSrc  = Join-Path $ScriptDir "manifest_local.xml"
$SharedFolder = "$env:USERPROFILE\Documents\WordAddins"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  文書校正アシスタント セットアップ" -ForegroundColor Cyan
Write-Host "  （Caddy ローカルホスト版）" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# ---- Step 1: office.js の確認 ----
$OfficeJs = Join-Path $ScriptDir "office.js"
if (-not (Test-Path $OfficeJs)) {
    Write-Host "⚠️  office.js が見つかりません。" -ForegroundColor Yellow
    $ans = Read-Host "今すぐダウンロードしますか？ [y/N]"
    if ($ans -eq "y" -or $ans -eq "Y") {
        & "$ScriptDir\fetch_office_js.ps1"
    } else {
        Write-Host "セットアップを中断しました。fetch_office_js.ps1 を実行してから再実行してください。"
        exit 1
    }
} else {
    Write-Host "✓ office.js を確認しました。" -ForegroundColor Green
}

# ---- Step 2: Caddy のインストール確認 ----
if (-not (Get-Command caddy -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "⚠️  Caddy がインストールされていません。" -ForegroundColor Yellow
    Write-Host "   winget install Caddy.Caddy を実行してからこのスクリプトを再実行してください。"
    exit 1
} else {
    $caddyVer = caddy version 2>&1 | Select-Object -First 1
    Write-Host "✓ Caddy を確認しました（$caddyVer）" -ForegroundColor Green
}

# ---- Step 3: CA 証明書の登録 ----
Write-Host ""
Write-Host "Caddy のローカル CA 証明書を OS に登録します..." -ForegroundColor Cyan
caddy trust
Write-Host "✓ CA 証明書を登録しました。" -ForegroundColor Green

# ---- Step 4: manifest_local.xml を共有フォルダに配置 ----
if (-not (Test-Path $SharedFolder)) {
    New-Item -ItemType Directory -Path $SharedFolder | Out-Null
}
Copy-Item $ManifestSrc "$SharedFolder\manifest_local.xml" -Force

Write-Host ""
Write-Host "✓ manifest_local.xml を $SharedFolder にコピーしました。" -ForegroundColor Green

# ---- Step 5: Caddy サービス登録 ----
Write-Host ""
$ans = Read-Host "Caddy を Windows サービスとして登録しますか？（PC 起動時に自動起動） [y/N]"
if ($ans -eq "y" -or $ans -eq "Y") {
    caddy service install
    caddy service start
    Write-Host "✓ Caddy サービスを登録・起動しました。" -ForegroundColor Green
} else {
    Write-Host "手動起動する場合: caddy run --config `"$ScriptDir\Caddyfile`""
}

# ---- Step 6: Word への登録案内 ----
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
Write-Host "   から「文書校正アシスタント（ローカル）」を選択"
Write-Host ""
