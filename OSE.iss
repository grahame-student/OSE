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
Source: DLLs\comcat.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: DLLs\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
Source: DLLs\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: DLLs\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: DLLs\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver

Source: ..\..\..\..\..\Windows\System32\MSCOMCTL.OCX; DestDir: {sys}; Flags: promptifolder regserver sharedfile
Source: OSE.exe; DestDir: {app}; Flags: promptifolder 32bit
Source: Factions.data; DestDir: {app}
Source: ..\..\..\..\..\Windows\System32\comdlg32.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: Spells.data; DestDir: {app}
Source: Items.data; DestDir: {app}

[Icons]
Name: {group}\OSE; Filename: {app}\OSE.exe; WorkingDir: {app}; IconFilename: {app}\OSE.exe; IconIndex: 0

[Dirs]
Name: {app}\Screenshot\

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
