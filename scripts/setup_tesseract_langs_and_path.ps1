param(
    [string]$TesseractDir = "C:\Program Files\Tesseract-OCR"
)

$ErrorActionPreference = "Stop"

function Add-UserPathIfMissing([string]$PathToAdd) {
    $current = [Environment]::GetEnvironmentVariable("Path", "User")
    $parts = @()
    if ($current) { $parts = $current -split ";" }
    if ($parts -notcontains $PathToAdd) {
        $newPath = ($parts + $PathToAdd | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Unique) -join ";"
        [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
        Write-Host "Added to USER PATH: $PathToAdd"
    } else {
        Write-Host "Already in USER PATH: $PathToAdd"
    }
    if (($env:Path -split ";") -notcontains $PathToAdd) {
        $env:Path = "$env:Path;$PathToAdd"
        Write-Host "Added to current session PATH: $PathToAdd"
    }
}

if (-not (Test-Path $TesseractDir)) {
    throw "Tesseract directory not found: $TesseractDir. Install Tesseract first."
}

$tesseractExe = Join-Path $TesseractDir "tesseract.exe"
if (-not (Test-Path $tesseractExe)) {
    throw "tesseract.exe not found in: $TesseractDir"
}

$tessdataDir = Join-Path $TesseractDir "tessdata"
New-Item -ItemType Directory -Force -Path $tessdataDir | Out-Null

$langs = @("eng", "hun", "slk", "slv", "hrv", "ces", "ron")
$base = "https://github.com/tesseract-ocr/tessdata_fast/raw/main"

foreach ($lang in $langs) {
    $dst = Join-Path $tessdataDir "$lang.traineddata"
    if (Test-Path $dst) {
        Write-Host "Exists: $lang.traineddata"
        continue
    }
    $url = "$base/$lang.traineddata"
    Write-Host "Downloading: $url"
    Invoke-WebRequest -Uri $url -OutFile $dst
}

Add-UserPathIfMissing -PathToAdd $TesseractDir

Write-Host ""
Write-Host "Installed language data in: $tessdataDir"
Write-Host "Verifying Tesseract..."
& $tesseractExe --version | Select-Object -First 1
& $tesseractExe --list-langs
