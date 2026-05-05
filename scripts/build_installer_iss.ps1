param(
    [string]$IsccPath = "",
    [string]$IssScript = "installer\\string_image_ocr.iss"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $IssScript)) {
    throw "ISS script not found: $IssScript"
}

if ([string]::IsNullOrWhiteSpace($IsccPath)) {
    $candidates = @(
        "$env:ProgramFiles(x86)\\Inno Setup 6\\ISCC.exe",
        "$env:ProgramFiles\\Inno Setup 6\\ISCC.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            $IsccPath = $c
            break
        }
    }
}

if ([string]::IsNullOrWhiteSpace($IsccPath) -or -not (Test-Path $IsccPath)) {
    throw "Inno Setup compiler (ISCC.exe) not found. Install Inno Setup 6 or pass -IsccPath."
}

& $IsccPath $IssScript
Write-Host "Installer build finished. Output directory: dist\\installer"
