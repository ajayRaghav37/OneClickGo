; Script generated by ANIco.in 2011 � All Rights Reserved

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
PrivilegesRequired=admin
AppId={{E1045EB2-8685-4D82-A8D7-97CBDA3B354A}
AppName=OneClick Go!
AppVersion=2.7
;AppVerName=OneClick Go! 2.7
AppPublisher=ANIco.in
AppPublisherURL=http://www.ANIco.in/
AppSupportURL=http://www.ANIco.in/
AppUpdatesURL=http://www.ANIco.in/
DefaultDirName={pf}\ANIco.in\OneClick Go!
DefaultGroupName=ANIco.in
AllowNoIcons=yes
LicenseFile=E:\OneClick Go! 2.7\Source Code\Program Files\Documents\END USER LICENSE AGREEMENT.rtf
OutputDir=E:\OneClick Go! 2.7
OutputBaseFilename=OneClick_Go_Final_Release_Setup
SetupIconFile=E:\OneClick Go! 2.7\Source Code\Visual Interface\OneClick Go!.ico
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 0,6.1

[Files]
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\COMDLG32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Setup Files\System Files\MCLhotkey.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\OneClick Go!.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Documents\END USER LICENSE AGREEMENT.rtf"; DestDir: "{app}\Documents"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Documents\Help.pdf"; DestDir: "{app}\Documents"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NextPressedPause.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NextPressedPlay.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NPSback.bmp"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NPSback1.bmp"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NPSback2.bmp"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NPSBlockerL.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\NPSBlockerR.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PauseDefault.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PausePressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayDefault.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayPauseADefault.bmp"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayPauseAPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayPauseLDefault.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayPauseLPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PlayPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PrevNextDefault.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PrevNextNPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PrevNextPPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PrevPressedPause.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\PrevPressedPlay.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\StopDefault.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\StopPressed.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\StopPressedPause.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
Source: "E:\OneClick Go! 2.7\Source Code\Program Files\Skin\StopPressedPlay.BMP"; DestDir: "{app}\Skin"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\OneClick Go!\OneClick Go! v2.7"; Filename: "{app}\OneClick Go!.exe"
Name: "{group}\OneClick Go!\{cm:UninstallProgram,OneClick Go!}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\OneClick Go!"; Filename: "{app}\OneClick Go!.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\OneClick Go!"; Filename: "{app}\OneClick Go!.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\OneClick Go!.exe"; Description: "{cm:LaunchProgram,OneClick Go!}"; Flags: nowait postinstall skipifsilent