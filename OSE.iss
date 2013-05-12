[Setup]
AppName=OSE v0.1.8
AppVerName=OSE v0.1.8
DefaultDirName={pf32}\OSE\
DefaultGroupName=Oblivion Save Editor
OutputDir=C:\Users\Grahame\Documents\VB6\OSE\Package\
AppID={{6EDB5158-B63E-466C-8976-47E75B99F350}
VersionInfoVersion=0.1.8
VersionInfoProductName=OSE
VersionInfoProductVersion=0.1.8
AppVersion=OSE v0.1.8
OutputBaseFilename=Install OSE 0.1.8
UsePreviousAppDir=true
UninstallDisplayIcon={app}\OSE.exe
UsePreviousGroup=false
SetupIconFile=C:\Users\Grahame\Documents\VB6\OSE\oblivion 48x48.ico

[Files]
; [Bootstrap Files]
; @COMCAT.DLL,$(WinSysPathSysFile),$(DLLSelfRegister),,5/31/98 12:00:00 AM,22288,4.71.1460.1
Source: DLLs\comcat.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @asycfilt.dll,$(WinSysPathSysFile),,,11/20/10 12:18:04 PM,67584,6.1.7601.17514
Source: DLLs\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; @olepro32.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,11/20/10 12:20:49 PM,90112,6.1.7601.17514
Source: DLLs\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @oleaut32.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,8/27/11 4:26:27 AM,571904,6.1.7601.17676
Source: DLLs\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
; @msvbvm60.dll,$(WinSysPathSysFile),$(DLLSelfRegister),,7/14/09 1:15:50 AM,1386496,6.0.98.15
Source: DLLs\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver

; [Setup1 Files]
; @MSCOMCTL.OCX,$(WinSysPath),$(DLLSelfRegister),$(Shared),6/6/12 7:59:42 PM,1070152,6.1.98.34
Source: ..\..\..\..\..\Windows\System32\MSCOMCTL.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
; @OSE.exe,$(AppPath),,,11/15/12 8:39:42 PM,143360,1.0.0.0
Source: OSE.exe; DestDir: {app}; Flags: promptifolder 32bit
Source: Factions.data; DestDir: {app}
Source: ..\..\..\..\..\Windows\System32\comdlg32.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: Spells.data; DestDir: {app}
Source: Items.data; DestDir: {app}

[Icons]
Name: {group}\OSE; Filename: {app}\OSE.exe; WorkingDir: {app}; IconFilename: {app}\OSE.exe; IconIndex: 0

[Dirs]
Name: {app}\Screenshots\

[Code]
/////////////////////////////////////////////////////////////////////
function GetUninstallString(): String;
var
  sUnInstPath: String;
  sUnInstallString: String;
begin
  sUnInstPath := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#emit SetupSetting("AppId")}_is1');
  sUnInstallString := '';
  if not RegQueryStringValue(HKLM, sUnInstPath, 'UninstallString', sUnInstallString) then
    RegQueryStringValue(HKCU, sUnInstPath, 'UninstallString', sUnInstallString);
  Result := sUnInstallString;
end;


/////////////////////////////////////////////////////////////////////
function IsUpgrade(): Boolean;
begin
  Result := (GetUninstallString() <> '');
end;


/////////////////////////////////////////////////////////////////////
function UnInstallOldVersion(): Integer;
var
  sUnInstallString: String;
  iResultCode: Integer;
begin
// Return Values:
// 1 - uninstall string is empty
// 2 - error executing the UnInstallString
// 3 - successfully executed the UnInstallString

  // default return value
  Result := 0;

  // get the uninstall string of the old app
  sUnInstallString := GetUninstallString();
  if sUnInstallString <> '' then begin
    sUnInstallString := RemoveQuotes(sUnInstallString);
    if Exec(sUnInstallString, '/SILENT /NORESTART /SUPPRESSMSGBOXES','', SW_HIDE, ewWaitUntilTerminated, iResultCode) then
      Result := 3
    else
      Result := 2;
  end else
    Result := 1;
end;

/////////////////////////////////////////////////////////////////////
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if (CurStep=ssInstall) then
  begin
    if (IsUpgrade()) then
    begin
      UnInstallOldVersion();
    end;
  end;
end;
