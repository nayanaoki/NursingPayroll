; Inno Setup スクリプト — 訪問看護給与（3職種共通・ユーザー単位インストール）
; 直接は使わず build_all.py から呼ぶ（職種ごとに /D で値を渡す）:
;   ISCC /DAppName=... /DAppExe=... /DAppIdStr=... /DAppVersion=... /DOutName=... installer.iss
;
; 管理者権限不要。インストール先(LocalAppData\Programs)は書込可能なので自動更新が動く。

#ifndef AppName
  #define AppName "訪問看護給与_看護師"
#endif
#ifndef AppExe
  #define AppExe "訪問看護給与_看護師.exe"
#endif
#ifndef AppIdStr
  #define AppIdStr "{{4C7B1E20-5A31-4D8E-9B02-NURSEPAY0NS1}"
#endif
#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef DistDir
  #define DistDir "dist"
#endif
#ifndef OutName
  #define OutName "NursingPayroll_NS_setup"
#endif

[Setup]
AppId={#AppIdStr}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppName}
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=no
PrivilegesRequired=lowest
OutputDir=release
OutputBaseFilename={#OutName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayName={#AppName}

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにショートカットを作成する"; GroupDescription: "追加アイコン:"

[Files]
Source: "{#DistDir}\{#AppExe}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{group}\{#AppName} をアンインストール"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "今すぐ起動する"; Flags: nowait postinstall skipifsilent

; アンインストールしても %APPDATA%\{#AppName}（データ）は残す（誤削除防止）。
