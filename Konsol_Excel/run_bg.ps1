# run_bg.ps1 - Background launcher with watchdog + auto-resume
param(
    [Parameter(Mandatory=$true)]
    [string]$File,
    [string]$Years = "2020-2025",
    [string]$Output,
    [string]$VerificationDate,
    [switch]$AttributeSearch,
    [int]$Concurrent = 5
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ScriptDir "venv\Scripts\python.exe"
$LogFile = Join-Path $ScriptDir "watchdog.log"
$PidFile = Join-Path $ScriptDir "app.pid"

$Python = if (Test-Path $VenvPython) { $VenvPython } else { "python" }

function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$Time] $Message"
    Write-Host $Line
    Add-Content -Path $LogFile -Value $Line
}

function Start-App {
    param([bool]$Resume = $false)

    $baseArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-Command"
    )

    Set-Location $ScriptDir

    $cmdArgs = @()
    $cmdArgs += "-f"
    $cmdArgs += "`"$File`""
    $cmdArgs += "-y"
    $cmdArgs += $Years
    if ($Output) { $cmdArgs += "-o"; $cmdArgs += "`"$Output`"" }
    if ($VerificationDate) { $cmdArgs += "--verification-date"; $cmdArgs += $VerificationDate }
    if ($AttributeSearch) { $cmdArgs += "--attribute-search" }
    if ($Resume) { $cmdArgs += "--resume" }
    $cmdArgs += "--concurrent"; $cmdArgs += $Concurrent

    $pythonCmd = "& `"$Python`" main.py $($cmdArgs -join ' ')"

    if ($Resume) {
        Write-Log "Restarting app with --resume..."
    } else {
        Write-Log "Starting application..."
    }

    $args = $baseArgs + $pythonCmd
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $args -PassThru -WindowStyle Minimized
    $proc.Id | Out-File -FilePath $PidFile -Force
    Write-Log "App started with PID: $($proc.Id)"
    return $proc
}

function Is-Running {
    param([int]$Pid)
    try {
        $proc = Get-Process -Id $Pid -ErrorAction Stop
        return (-not $proc.HasExited)
    } catch {
        return $false
    }
}

function Stop-App {
    if (Test-Path $PidFile) {
        $Pid = [int](Get-Content $PidFile)
        try {
            $proc = Get-Process -Id $Pid -ErrorAction Stop
            if (-not $proc.HasExited) {
                Write-Log "Stopping app (PID: $Pid)..."
                $proc.Kill()
                Start-Sleep -Seconds 2
            }
        } catch { }
    }
}

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "FGIS ARSHIN - Background Watchdog v1.0 (auto-resume)" -ForegroundColor Yellow
Write-Host "======================================================================" -ForegroundColor Cyan

if ($args[0] -eq "--stop") {
    Write-Log "Stop signal received"
    Stop-App
    Write-Host "Application stopped" -ForegroundColor Green
    exit 0
}

if ($args[0] -eq "--status") {
    if (Test-Path $PidFile) {
        $Pid = [int](Get-Content $PidFile)
        if (Is-Running $Pid) {
            Write-Host "App is RUNNING (PID: $Pid)" -ForegroundColor Green
        } else {
            Write-Host "App is STOPPED (last PID: $Pid)" -ForegroundColor Red
        }
    } else {
        Write-Host "App is STOPPED (no PID file)" -ForegroundColor Red
    }
    exit 0
}

if (-not $File) {
    Write-Host "Usage: .\run_bg.ps1 -File запрос.xlsx [-Years 2020-2025] [-Output результат.csv]" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  .\run_bg.ps1 --status    Check if app is running"
    Write-Host "  .\run_bg.ps1 --stop      Stop the watchdog and app"
    exit 1
}

Set-Location $ScriptDir

$AppProc = $null
$IsRestart = $false

while ($true) {
    if ($AppProc -and $AppProc.HasExited) {
        Write-Log "App exited with code: $($AppProc.ExitCode)"

        if ($AppProc.ExitCode -eq 0) {
            Write-Log "App finished normally. Exiting watchdog."
            Remove-Item $PidFile -ErrorAction SilentlyContinue
            break
        }

        $IsRestart = $true
        Write-Log "App crashed. Restarting with --resume in 5 seconds..."
        Start-Sleep -Seconds 5
        $AppProc = $null
    }

    if (-not $AppProc) {
        $AppProc = Start-App -Resume $IsRestart
    }

    Start-Sleep -Seconds 10

    if ($AppProc) {
        try { $AppProc.Refresh() } catch { }
        if ($AppProc.HasExited) {
            continue
        }
    }
}

Write-Log "Watchdog exiting"
