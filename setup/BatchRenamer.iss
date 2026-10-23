#define MyAppName "批量重命名工具"
#define MyAppVersion "1.1.0"
#define MyAppPublisher "xfastres"
#define MyAppExeName "BatchRenamer.exe"

[Setup]
AppId={{A7C3E5B1-4D2F-4A9E-8C61-9B2E7F4D1A08}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
DisableWelcomePage=yes
DisableReadyPage=yes
AllowNoIcons=yes
OutputDir=..\dist
OutputBaseFilename=BatchRenamer-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
MinVersion=10.0
ShowLanguageDialog=no
UninstallDisplayName={#MyAppName}
VersionInfoVersion={#MyAppVersion}
VersionInfoProductName={#MyAppName}
SetupIconFile=..\app.ico
UninstallDisplayIcon={app}\app.ico
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
SetupAppTitle=安装
SetupWindowTitle=安装 {#MyAppName}
ButtonBack=< 上一步
ButtonNext=下一步 >
ButtonInstall=安装
ButtonOK=确定
ButtonCancel=取消
ButtonFinish=完成
ButtonBrowse=浏览...
BrowseDialogTitle=浏览文件夹
BrowseDialogLabel=选择一个文件夹，然后点击「确定」。
WizardSelectDir=选择安装位置
SelectDirDesc=安装程序会把 [name] 装到下面的文件夹。
SelectDirLabel3=安装到这个文件夹：
SelectDirBrowseLabel=要换一个位置，请点击「浏览」。
DiskSpaceGBLabel=至少需要 [gb] GB 可用空间。
DiskSpaceMBLabel=至少需要 [mb] MB 可用空间。
WizardSelectTasks=附加选项
SelectTasksDesc=还可以同时完成这些事项。
SelectTasksLabel2=请选择要执行的附加任务，然后点击「安装」。
WizardInstalling=正在安装
InstallingLabel=正在安装 [name]，请稍候。
FinishedHeadingLabel=已经安装完成
FinishedLabel=[name] 已安装完成。可以用桌面或开始菜单中的快捷方式打开。
FinishedLabelNoIcons=[name] 已安装完成。
ClickFinish=点击「完成」退出安装程序。
ExitSetupTitle=退出安装
ExitSetupMessage=安装还没完成。现在退出的话，程序将不会被安装。确定要退出吗？
StatusExtractFiles=正在复制文件...
StatusCreateIcons=正在创建快捷方式...

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; GroupDescription: "附加任务："; Flags: checkedonce

[Files]
Source: "..\BatchRenamer.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\BatchRenamer.Core.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\BatchRenamerComposer.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\启动.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\启动.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\操作说明.md"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\BatchRenamer.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Comment: "{#MyAppName}"; AppUserModelID: "xfastres.BatchRenamer"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; IconFilename: "{app}\app.ico"; Comment: "{#MyAppName}"; AppUserModelID: "xfastres.BatchRenamer"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即运行 {#MyAppName}"; Flags: nowait postinstall skipifsilent shellexec
