; Iteration 88: AhkWord's installer (Inno Setup 6), made by make_installer.ps1, which passes
;   /DIteration=NN /DSourceDir=<the staged program folder> /DOutputDir=<..\AhkWord-releases> /DRepo=<this repo>
; Installed for the current Windows user only (no administrator, no UAC prompt), in %LOCALAPPDATA%\Programs\AhkWord:
; AhkWord.exe (self-contained: .NET inside), AhkWord's own files beside it, and AutoHotkey64.exe for the hotkeys,
; snippets and macros. Settings, templates and recovery copies stay in %APPDATA%\AhkWord: uninstalling leaves them.

#ifndef Iteration
  #error Run make_installer.ps1 (it passes /DIteration, /DSourceDir, /DOutputDir and /DRepo).
#endif

[Setup]
; (the AppId must never change: Windows knows the installed AhkWord by it, and a newer installer replaces it)
AppId={{6F0A3C52-8B7E-4D1A-9C6E-2B5D7A4E1F88}
AppName=AhkWord
AppVersion={#Iteration}
AppVerName=AhkWord {#Iteration} (installer preview)
AppPublisher=burque505
AppPublisherURL=https://github.com/burque505/AhkWord
AppSupportURL=https://github.com/burque505/AhkWord/issues
AppUpdatesURL=https://github.com/burque505/AhkWord/releases
VersionInfoVersion={#Iteration}.0.0.0
PrivilegesRequired=lowest
DefaultDirName={autopf}\AhkWord
DefaultGroupName=AhkWord
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
LicenseFile={#Repo}\LICENSE
SetupIconFile={#Repo}\images\icon\ahkword.ico
UninstallDisplayIcon={app}\AhkWord.exe
OutputDir={#OutputDir}
OutputBaseFilename=AhkWord-{#Iteration}-setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no
ChangesAssociations=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "es"; MessagesFile: "compiler:Languages\Spanish.isl"

[CustomMessages]
en.OpenWith=Offer AhkWord in "Open with" for .rtf, .docx, .odt, .txt and .html files (doesn't change your default programs)
es.OpenWith=Ofrecer AhkWord en "Abrir con" para archivos .rtf, .docx, .odt, .txt y .html (no cambia sus programas predeterminados)
en.RunNow=Start AhkWord
es.RunNow=Iniciar AhkWord

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "openwith"; Description: "{cm:OpenWith}"; Flags: unchecked

[InstallDelete]
; (an older AhkWord's libraries and .NET: whatever this one doesn't carry any more mustn't stay)
Type: files; Name: "{app}\*.dll"
Type: files; Name: "{app}\*.deps.json"
Type: files; Name: "{app}\*.runtimeconfig.json"
Type: filesandordirs; Name: "{app}\runtimes"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\AhkWord"; Filename: "{app}\AhkWord.exe"; WorkingDir: "{app}"
Name: "{autodesktop}\AhkWord"; Filename: "{app}\AhkWord.exe"; WorkingDir: "{app}"; Tasks: desktopicon

[Registry]
; "Open with" only (HKCU, this user): Windows offers AhkWord for these types, the default program stays as it is
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe"; ValueType: string; ValueName: "FriendlyAppName"; ValueData: "AhkWord"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\DefaultIcon"; ValueType: string; ValueData: """{app}\AhkWord.exe"",0"; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\shell\open\command"; ValueType: string; ValueData: """{app}\AhkWord.exe"" ""%1"""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\SupportedTypes"; ValueType: string; ValueName: ".rtf"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\SupportedTypes"; ValueType: string; ValueName: ".docx"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\SupportedTypes"; ValueType: string; ValueName: ".odt"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\SupportedTypes"; ValueType: string; ValueName: ".txt"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\Applications\AhkWord.exe\SupportedTypes"; ValueType: string; ValueName: ".html"; ValueData: ""; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.rtf\OpenWithList\AhkWord.exe"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.docx\OpenWithList\AhkWord.exe"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.odt\OpenWithList\AhkWord.exe"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.txt\OpenWithList\AhkWord.exe"; Flags: uninsdeletekey; Tasks: openwith
Root: HKCU; Subkey: "Software\Classes\.html\OpenWithList\AhkWord.exe"; Flags: uninsdeletekey; Tasks: openwith

[Run]
Filename: "{app}\AhkWord.exe"; Description: "{cm:RunNow}"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent

