; Inno Setup script for String Image OCR Report GUI
#define AppName "String Image OCR Report"
#define AppVersion "2.2.0"
#define AppPublisher "sinkokrisztian"
#define AppExeName "StringImageOCRReport.exe"

[Setup]
AppId={{9B2D6E1F-1B02-4F11-8A39-9E7F6E6B7A4C}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\StringImageOCRReport
DefaultGroupName=String Image OCR Report
DisableProgramGroupPage=yes
OutputDir=dist\installer
OutputBaseFilename=StringImageOCRReport_Setup_{#AppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\nuitka\StringImageOCRReport.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\String Image OCR Report"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\String Image OCR Report"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,String Image OCR Report}"; Flags: nowait postinstall skipifsilent
