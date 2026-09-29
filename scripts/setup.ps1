# セットアップ（Windows）: .venv を作り、このツールをインストールする
#   .\setup.cmd                       （ダブルクリックでも可）
#   .\scripts\setup.ps1 -Dev          （開発者: テスト・lint も入れる）
param([switch]$Dev)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Repo = Split-Path $PSScriptRoot -Parent
Set-Location $Repo

# 1. Python 3.10 以上を探す
$py = $null; $pyArgs = @()
foreach ($c in @("py", "python", "python3")) {
    if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; if ($c -eq "py") { $pyArgs = @("-3") }; break }
}
if (-not $py) {
    Write-Host "Python が見つかりません。Python 3.10 以上をインストールしてください: https://www.python.org/downloads/" -ForegroundColor Red
    Write-Host "（インストーラーで「Add python.exe to PATH」にチェック）"
    exit 1
}
& $py @pyArgs -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)"
if ($LASTEXITCODE -ne 0) { Write-Host "Python 3.10 以上が必要です" -ForegroundColor Red; exit 1 }

# 2. 仮想環境
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "== 仮想環境 .venv を作成"
    & $py @pyArgs -m venv .venv
}
$vpy = Join-Path $Repo ".venv\Scripts\python.exe"
& $vpy -m pip install --upgrade pip --quiet

# 3. 本体（隣に devtools-common があればそれを使う）
Write-Host "== doc-to-markdown をインストール"
if ($Dev) { & $vpy scripts\install_deps.py --dev } else { & $vpy scripts\install_deps.py }
if ($LASTEXITCODE -ne 0) { exit 1 }

# 4. 確認
& ".venv\Scripts\doc2md.exe" --check-env
Write-Host ""
Write-Host "完了しました。" -ForegroundColor Green
Write-Host "  - VS Code でこのフォルダを開き、Copilot Chat で / を入力するとプロンプトが使えます"
Write-Host "  - ターミナルで使う場合: .\.venv\Scripts\Activate.ps1 を実行してから doc2md --help"
