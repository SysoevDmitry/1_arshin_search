# run.ps1
param(
    [string]$File,
    [string]$Years = "2020-2025",
    [string]$Output,
    [string]$VerificationDate,
    [switch]$AttributeSearch,
    [switch]$Resume,
    [int]$Concurrent = 5,
    [string]$Mirror = ""  # e.g. "https://pypi.tuna.tsinghua.edu.cn/simple"
)

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "FGIS ARSHIN - v3.0" -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Cyan

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ScriptDir "venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$Python = $null

$PipCommonArgs = @("--default-timeout=120", "--retries", "5")
$PipIndexArgs = @()
if ($Mirror) {
    $PipIndexArgs = @("-i", $Mirror, "--trusted-host", ($Mirror -replace '^https?://', '' -replace '/.*$', ''))
    Write-Host "Mirror: $Mirror" -ForegroundColor DarkGray
}

function Install-PipPackage {
    param([string[]]$Packages)
    Write-Host "Installing: $($Packages -join ', ')..." -ForegroundColor Yellow

    # Try 1: default PyPI
    if (-not $Mirror) {
        & $Python -m pip install @PipCommonArgs $Packages 2>&1
        if ($LASTEXITCODE -eq 0) { return $true }
        Write-Host "Default PyPI failed, trying mirrors..." -ForegroundColor Yellow
    }

    # Try 2: mirrors
    $mirrors = @(
        @("https://pypi.tuna.tsinghua.edu.cn/simple", "pypi.tuna.tsinghua.edu.cn"),
        @("https://mirrors.aliyun.com/pypi/simple", "mirrors.aliyun.com"),
        @("https://pypi.mirrors.ustc.edu.cn/simple", "pypi.mirrors.ustc.edu.cn")
    )
    foreach ($m in $mirrors) {
        $MirrorUrl = $m[0]
        $MirrorHost = $m[1]
        Write-Host "  Trying $MirrorUrl ..." -ForegroundColor DarkGray
        & $Python -m pip install @PipCommonArgs -i $MirrorUrl --trusted-host $MirrorHost $Packages 2>&1
        if ($LASTEXITCODE -eq 0) { return $true }
    }

    return $false
}

Write-Host "`nChecking Python..." -ForegroundColor Cyan
$PythonExe = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $ver = & $cmd --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            $PythonExe = $cmd
            Write-Host "Found: $ver" -ForegroundColor Green
            break
        }
    } catch { }
}

if (-not $PythonExe) {
    Write-Host "ERROR: Python not found!" -ForegroundColor Red
    Write-Host "Install Python 3.10+ from https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

if (Test-Path $VenvPython) {
    Write-Host "Virtual environment found" -ForegroundColor Green
    $Python = $VenvPython
} else {
    Write-Host "Creating virtual environment..." -ForegroundColor Yellow
    & $PythonExe -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to create virtual environment" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Virtual environment created" -ForegroundColor Green
    $Python = $VenvPython
}

Write-Host "`nChecking pip..." -ForegroundColor Cyan
$PipVersion = & $Python -m pip --version 2>&1
Write-Host "Current: $PipVersion" -ForegroundColor Green

Write-Host "`nChecking dependencies..." -ForegroundColor Cyan
$CheckResult = & $Python -c "import pandas, openpyxl, aiohttp, requests, ttkbootstrap, tqdm" 2>&1
if ($LASTEXITCODE -ne 0) {
    if (Install-PipPackage -Packages @("pandas", "openpyxl", "aiohttp", "requests", "ttkbootstrap", "tqdm")) {
        Write-Host "Dependencies OK" -ForegroundColor Green
    } else {
        Write-Host "Failed to install dependencies" -ForegroundColor Red
        Write-Host ""
        Write-Host "Tips:" -ForegroundColor Yellow
        Write-Host "  1. Check internet connection / firewall / VPN" -ForegroundColor Yellow
        Write-Host "  2. If in China, use: .\run.ps1 -Mirror https://pypi.tuna.tsinghua.edu.cn/simple" -ForegroundColor Yellow
        Write-Host "  3. Manually: $Python -m pip install --default-timeout=120 pandas openpyxl aiohttp requests ttkbootstrap" -ForegroundColor Yellow
        Read-Host "`nPress Enter to exit"
        exit 1
    }
} else {
    Write-Host "Dependencies OK" -ForegroundColor Green
}

Write-Host "`nStarting application..." -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""

Set-Location $ScriptDir

$Args = @()
if ($File) { $Args += "-f", $File }
if ($Years) { $Args += "-y", $Years }
if ($Output) { $Args += "-o", $Output }
if ($VerificationDate) { $Args += "--verification-date", $VerificationDate }
if ($AttributeSearch) { $Args += "--attribute-search" }
if ($Resume) { $Args += "--resume" }
if ($Concurrent) { $Args += "--concurrent", $Concurrent }

if (-not $File) {
    Write-Host "Usage: .\run.ps1 -File data.xlsx [-Years 2020-2025] [-Output results.csv] [-Resume] [-AttributeSearch]" -ForegroundColor Yellow
    Write-Host "       .\run.ps1 -File data.xlsx -Mirror https://pypi.tuna.tsinghua.edu.cn/simple" -ForegroundColor Yellow
    Write-Host "`nInteractive mode:" -ForegroundColor Cyan
    & $Python main.py -h
} else {
    & $Python main.py @Args
}

Read-Host "`nPress Enter to exit"
