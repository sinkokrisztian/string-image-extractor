param(
    [string]$PythonExe = ".venv\\Scripts\\python.exe",
    [string]$OutDir = "dist\\nuitka",
    [string]$ExeName = "StringImageOCRReport"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $PythonExe)) {
    throw "Python executable not found: $PythonExe"
}

& $PythonExe -m pip install --upgrade pip | Out-Host
& $PythonExe -m pip install nuitka ordered-set zstandard | Out-Host

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

& $PythonExe -m nuitka `
    --onefile `
    --assume-yes-for-downloads `
    --enable-plugin=tk-inter `
    --windows-console-mode=disable `
    --output-dir=$OutDir `
    --output-filename="$ExeName.exe" `
    src\\ocr_report_gui.py

Write-Host "Nuitka build finished. Output directory: $OutDir"
