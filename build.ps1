# White Border Cropper build script
# Usage: .\build.ps1

param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$AppName = "cropper"
$AppVersion = "0.3"
$ExeName = "$AppName-v$AppVersion"

function Invoke-Checked {
    param(
        [string]$FilePath,
        [string[]]$Arguments
    )

    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed ($LASTEXITCODE): $FilePath $($Arguments -join ' ')"
    }
}

function Remove-FileWithRetry {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        return
    }

    for ($i = 1; $i -le 5; $i++) {
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            return
        }
        catch {
            if ($i -eq 5) {
                throw "Cannot overwrite $Path. Close the app if it is running, then run build again."
            }
            Start-Sleep -Milliseconds (500 * $i)
        }
    }
}

function Assert-OutputNotRunning {
    param([string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fileName = [System.IO.Path]::GetFileName($fullPath).Replace("'", "''")
    $running = Get-CimInstance Win32_Process -Filter "Name = '$fileName'" |
        Where-Object { $_.ExecutablePath -eq $fullPath }

    if ($running) {
        $pids = ($running | ForEach-Object { $_.ProcessId }) -join ", "
        throw "Output exe is running (PID: $pids). Close it, then run build again."
    }
}

function New-Venv {
    param([string]$Path)

    Write-Host "[INFO] Creating virtual environment: $Path" -ForegroundColor Cyan
    if (Get-Command python -ErrorAction SilentlyContinue) {
        Invoke-Checked "python" @("-m", "venv", $Path)
    }
    elseif (Get-Command py -ErrorAction SilentlyContinue) {
        Invoke-Checked "py" @("-3", "-m", "venv", $Path)
    }
    else {
        throw "Python 3 was not found. Please install Python first."
    }
}

Push-Location $PSScriptRoot
try {
    $venvPython = Join-Path $PSScriptRoot "venv_clean\Scripts\python.exe"
    if (-not (Test-Path $venvPython)) {
        $venvDir = Join-Path $PSScriptRoot ".venv"
        $venvPython = Join-Path $venvDir "Scripts\python.exe"
        if (-not (Test-Path $venvPython)) {
            New-Venv $venvDir
        }
    }

    New-Item -ItemType Directory -Force -Path "build" | Out-Null
    New-Item -ItemType Directory -Force -Path "dist" | Out-Null
    $outputExe = Join-Path $PSScriptRoot "dist\$ExeName.exe"
    Assert-OutputNotRunning $outputExe

    if (-not $SkipInstall) {
        Write-Host "[INFO] Installing dependencies..." -ForegroundColor Cyan
        Invoke-Checked $venvPython @(
            "-m", "pip", "install",
            "--disable-pip-version-check",
            "--quiet",
            "-r", "requirement.txt",
            "pyinstaller>=6.0"
        )
    }

    $pyBase = & $venvPython -c "import sys; print(sys.base_prefix)"
    $tclBin = Join-Path $pyBase "Library\bin"
    if (Test-Path $tclBin) {
        $env:Path = "$tclBin;$env:Path"
    }

    Remove-FileWithRetry $outputExe

    Write-Host "[INFO] Building $ExeName.exe ..." -ForegroundColor Cyan
    Invoke-Checked $venvPython @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--log-level", "WARN",
        "--onefile",
        "--windowed",
        "--noupx",
        "--name", $ExeName,
        "--distpath", "dist",
        "--workpath", "build\pyinstaller",
        "--specpath", "build",
        "--collect-all", "pymupdf",
        "--collect-submodules", "PIL",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "main.py"
    )

    Write-Host "[DONE] Output: dist\$ExeName.exe" -ForegroundColor Green
}
finally {
    Pop-Location
}
