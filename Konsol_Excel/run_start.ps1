# run_start.ps1 - Quick start, no checks
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

$Python = if (Test-Path "venv\Scripts\python.exe") { "venv\Scripts\python.exe" } else { "python" }

Write-Host "Starting..." -ForegroundColor Cyan
& $Python main.py -f Запрос.xlsx -y 2010-2026 --attribute-search $args

Read-Host "`nPress Enter to exit"
