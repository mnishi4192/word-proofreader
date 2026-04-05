# ============================================================
# office.js をローカルに取得するスクリプト（Windows 用）
# ============================================================
# Microsoft の CDN から office.js をダウンロードし、
# このフォルダに保存します。
# 一度だけ実行すれば OK です（以降はインターネット不要）。
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutFile   = Join-Path $ScriptDir "office.js"
$Url       = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  office.js ダウンロード" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "取得先: $Url"
Write-Host "保存先: $OutFile"
Write-Host ""

if (Test-Path $OutFile) {
    $ans = Read-Host "既に office.js が存在します。上書きしますか？ [y/N]"
    if ($ans -ne "y" -and $ans -ne "Y") {
        Write-Host "スキップしました。"
        exit 0
    }
}

try {
    Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
    $size = (Get-Item $OutFile).Length
    Write-Host ""
    Write-Host "✓ office.js を取得しました（$size バイト）" -ForegroundColor Green
    Write-Host ""
    Write-Host "次のステップ:" -ForegroundColor Cyan
    Write-Host "  caddy trust                          # CA 証明書を OS に登録（初回のみ、管理者権限）"
    Write-Host "  caddy run --config Caddyfile         # Caddy を起動"
    Write-Host ""
} catch {
    Write-Host "エラー: office.js の取得に失敗しました。" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit 1
}
