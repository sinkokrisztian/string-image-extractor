# Installer Build (Windows)

This project can be packaged as a Windows `.exe` with Nuitka, then wrapped into a setup installer with Inno Setup.

## 1) Build GUI executable with Nuitka

From repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_nuitka_gui.ps1
```

Expected output:

- `dist\nuitka\StringImageOCRReport.exe`

## 2) Build installer with Inno Setup (ISCC)

Install **Inno Setup 6** first (includes `ISCC.exe`), then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer_iss.ps1
```

Expected output:

- `dist\installer\StringImageOCRReport_Setup_2.3.0.exe`

If `ISCC.exe` is not in a default location, pass explicit path:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_installer_iss.ps1 -IsccPath "C:\Path\To\ISCC.exe"
```
