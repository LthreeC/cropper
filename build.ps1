# White Border Cropper build script
# Usage: .\build.ps1

param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$AppName = "cropper"
$AppVersion = "0.5.1"
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
    $processName = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    $running = Get-Process -Name $processName -ErrorAction SilentlyContinue |
        Where-Object {
            try {
                $_.Path -eq $fullPath
            }
            catch {
                $false
            }
        }

    if ($running) {
        $pids = ($running | ForEach-Object { $_.Id }) -join ", "
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
    # Pillow's standard hook already covers supported image plugins.
    # Exclude large optional components that this application does not use.
    $pyInstallerArguments = @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--log-level", "WARN",
        "--onefile",
        "--windowed",
        "--noupx",
        "--name", $ExeName,
        "--distpath", "dist",
        "--workpath", "build\pyinstaller-compact",
        "--specpath", "build",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--exclude-module", "PIL.AvifImagePlugin",
        "--exclude-module", "PIL._avif",
        "--exclude-module", "PIL.ImageCms",
        "--exclude-module", "PIL._imagingcms",
        "--exclude-module", "PIL._imagingft",
        "--exclude-module", "PIL.ImageMorph",
        "--exclude-module", "PIL._imagingmorph",
        "--exclude-module", "win32ui",
        "--exclude-module", "Pythonwin",
        "--exclude-module", "numpy.random",
        "--exclude-module", "numpy.fft",
        "--exclude-module", "numpy.polynomial",
        "--exclude-module", "numpy.testing",
        "--exclude-module", "numpy.f2py",
        "--exclude-module", "numpy.ma",
        "--exclude-module", "numpy.typing",
        "--exclude-module", "numpy.matlib",
        "--exclude-module", "numpy.ctypeslib",
        "--exclude-module", "ssl",
        "--exclude-module", "_ssl",
        "--exclude-module", "_hashlib",
        "main.py"
    )
    Invoke-Checked $venvPython $pyInstallerArguments

    Write-Host "[DONE] Output: dist\$ExeName.exe" -ForegroundColor Green
}
finally {
    Pop-Location
}
