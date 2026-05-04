@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0"

set "ROOT=%cd%"
set "SCRIPT=src\image_ocr_match_report.py"

if not exist "%SCRIPT%" (
  echo ERROR: %SCRIPT% not found in %ROOT%
  exit /b 1
)

echo.
echo ============================================
echo   OCR Match Report Launcher
echo ============================================
echo Root: %ROOT%
echo.
echo Choose target language:
echo   1. Hungarian (hu) - OCR hun
echo   2. Slovak    (sk) - OCR slk
echo   3. Slovenian (sl) - OCR slv
echo   4. Croatian  (hr) - OCR hrv
echo   5. Czech     (cz) - OCR ces
echo   6. Romanian  (ro) - OCR ron
echo.
set /p CHOICE=Enter number [1-6]: 

set "TARGET_LANG="
set "TARGET_OCR="

if "%CHOICE%"=="1" set "TARGET_LANG=hu" & set "TARGET_OCR=hun"
if "%CHOICE%"=="2" set "TARGET_LANG=sk" & set "TARGET_OCR=slk"
if "%CHOICE%"=="3" set "TARGET_LANG=sl" & set "TARGET_OCR=slv"
if "%CHOICE%"=="4" set "TARGET_LANG=hr" & set "TARGET_OCR=hrv"
if "%CHOICE%"=="5" set "TARGET_LANG=cz" & set "TARGET_OCR=ces"
if "%CHOICE%"=="6" set "TARGET_LANG=ro" & set "TARGET_OCR=ron"

if "%TARGET_LANG%"=="" (
  echo Invalid choice.
  exit /b 1
)

set "DEFAULT_OUT=report_%TARGET_LANG%.xlsx"
set /p OUTFILE=Output Excel file [%DEFAULT_OUT%]: 
if "%OUTFILE%"=="" set "OUTFILE=%DEFAULT_OUT%"

set "DEFAULT_LOG=report_%TARGET_LANG%.log"
set /p LOGFILE=Log file [%DEFAULT_LOG%]: 
if "%LOGFILE%"=="" set "LOGFILE=%DEFAULT_LOG%"

set /p TESS_CMD=Optional full path to tesseract.exe (leave empty if in PATH): 

echo.
echo Running report for lang=%TARGET_LANG% (OCR=%TARGET_OCR%)...
echo.

if "%TESS_CMD%"=="" (
  python "%SCRIPT%" ^
    --root "%ROOT%" ^
    --target-lang "%TARGET_LANG%" ^
    --source-prefix "enis" ^
    --source-ocr-lang "eng" ^
    --target-ocr-lang "%TARGET_OCR%" ^
    --output "%OUTFILE%" ^
    --log-file "%LOGFILE%" ^
    --use-temp-local-copy ^
    --verbose
) else (
  python "%SCRIPT%" ^
    --root "%ROOT%" ^
    --target-lang "%TARGET_LANG%" ^
    --source-prefix "enis" ^
    --source-ocr-lang "eng" ^
    --target-ocr-lang "%TARGET_OCR%" ^
    --tesseract-cmd "%TESS_CMD%" ^
    --output "%OUTFILE%" ^
    --log-file "%LOGFILE%" ^
    --use-temp-local-copy ^
    --verbose
)

if errorlevel 1 (
  echo.
  echo Run failed. Check "%LOGFILE%" for details.
  exit /b 1
)

echo.
echo Done. Output: %OUTFILE%
echo Log: %LOGFILE%
exit /b 0


