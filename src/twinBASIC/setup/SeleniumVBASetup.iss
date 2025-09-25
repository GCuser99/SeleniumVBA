// With this installer, the user can (re)install the DLL in a location
// of their choice (DisableDirPage=no). If the user had already
// installed an older (or even same or newer) version compared to the one
// user is currently installing, and then decides to change the
// install location, this installer (with help from UninsIS) will first
// uninstall the previously installed version from the old location,
// so that there is only one version on the system at any one point in time. 
#define AppName "SeleniumVBA"
#define AppGUID "{38ED0FFA-E3F3-41C4-B601-B8B2BC442ACD}"
#define AppPublisher "GCUser99"
#define AppURL "https://github.com/GCuser99/SeleniumVBA"
#define AppHelpURL "https://github.com/GCuser99/SeleniumVBA/wiki"
#define InstallerName "SeleniumVBADLLSetup"
#define DLL64FilePath "..\Build\SeleniumVBA_win64.dll"
#define DLL32FilePath "..\Build\SeleniumVBA_win32.dll"
#define LicenseFilePath "..\..\..\LICENSE.txt"
#define TestFolderPath ".\test_documents"
#define UtilitiesPath "..\..\Utilities"
#define LogoFilePath ".\logo_setup.bmp"
#define RequirementsFilePath ".\readme.rtf"
#define SetupOutputFolderPath "..\..\..\dist" 
#define AppVersion GetVersionNumbersString(DLL64FilePath)

[Setup]
AppId={{#AppGUID}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
; Set default install location
DefaultDirName={localappdata}\{#AppName}
DefaultGroupName={#AppName}
; Remove the following line to run in administrative
; install mode (install for all users.)
PrivilegesRequired=lowest
OutputBaseFilename={#InstallerName}
LicenseFile={#LicenseFilePath}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
; Uninstallable determines if Inno Setup's 
; automatic uninstaller is to be included in
; the installation folder - this must be set to
; "yes" for the PrepareToInstall code to function
; correctly
Uninstallable=yes
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
WizardImageFile={#LogoFilePath}
DisableWelcomePage=no
DisableProgramGroupPage=yes
InfoBeforeFile={#RequirementsFilePath}
; DisableDirPage must be set to "no" to allow 
; User to select a different install location
; if updating
DisableDirPage=no
OutputDir={#SetupOutputFolderPath}

[Languages]

Name: "english"; MessagesFile: "compiler:Default.isl"

[Components]

Name: "pkg_core"; Description: "SeleniumVBA ActiveX Dll"; Types: full compact custom; Flags: fixed;
Name: "pkg_utils";  Description: "PowerShell Utilities"; Types: full compact custom; Flags: fixed;
Name: "pkg_docs";  Description: "MS Excel and Access Test Documents"; Types: full compact custom;
  
[Messages]

// WelcomeLabel2=This will install [name/ver] on your computer.%n%nIt is recommended that you close all other applications before continuing.
FinishedLabel=Setup has finished installing [name] on your computer. A shortcut to the DLL folder can be found on your Desktop.
ClickFinish=If left checked, the optional task(s) below will be performed. Click Finish to complete and exit Setup.

[Files]
Source: {#DLL64FilePath}; DestDir: {app};  Flags: ignoreversion regserver ; Check: InstallX64; Components: pkg_core;
Source: {#DLL32FilePath}; DestDir: {app};  Flags: ignoreversion regserver ; Check: InstallX32; Components: pkg_core;
; Source: {#INIFilePath}; DestDir: {app};  Flags: ignoreversion uninsneveruninstall onlyifdoesntexist; Check: IsWin64; Components: pkg_core;
Source: {#TestFolderPath}\SeleniumVBA test subs for tB.xlsm; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs; 
Source: {#TestFolderPath}\SeleniumVBA test subs for tB.accdb; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#TestFolderPath}\readme.txt; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#UtilitiesPath}\cleanup_drivers.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\create_update_ini_file.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\launch_chrome_in_debugger_mode.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\launch_edge_in_debugger_mode.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\analize_registry.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#LicenseFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
Source: {#RequirementsFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
; Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme

[Icons]
Name: "{autodesktop}\SeleniumVBA - Shortcut"; Filename: "{app}"
Name: "{app}\wiki help documentation"; Filename: "{#AppHelpURL}"

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\utilities\create_update_ini_file.ps1"" -keepExistingValues -iniFilePath ""{app}\SeleniumVBA.ini"""; Description: "Create/update optional INI Settings file (old values will be preserved)"; Flags: postinstall skipifsilent runhidden;

[Registry]
;Add excel and access trusted location for examples
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVersion|Excel}.0\Excel\Security\Trusted Locations\{#AppName}"; ValueName: "Path"; ValueType: String; ValueData: "{app}\examples";  Flags: uninsdeletekey; Check: HasExcel;
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVersion|Access}.0\Access\Security\Trusted Locations\{#AppName}"; ValueName: "Path"; ValueType: String; ValueData: "{app}\examples";  Flags: uninsdeletekey; Check: HasAccess;
;Enable WScript host in case it's been disabled
;Root: HKCU; Subkey: "Software\Microsoft\Windows Script Host\Settings"; ValueName: "Enabled"; ValueType: dword; ValueData: 1; 

[Code]
const
  APPPATHSKEY = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths';
  OFFICE_UNKNOWN_BIT = -1;
  OFFICE_32_BIT = 0;
  OFFICE_64_BIT = 6;

// API call to determine bitness of an executable
function GetBinaryType(ApplicationName: string; var BinaryType: Integer): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

// The following code from Martin Prikryl - see
// https://stackoverflow.com/questions/47431008/getting-the-version-and-platform-of-office-application-from-windows-registry#47443674
function GetOfficeBitness(): Integer;
var
  officeApps: array[1..2] of string;
  i: Integer;
  officeAppPath: string;
  binaryType: Integer;
  keyFound: Boolean;
begin
  // Find the Office app binary path from Registry -
  //first try Excel, and if that fails then Access, then Word, etc 
  //once path is found, then use API to determine bitness of app
  Result := OFFICE_UNKNOWN_BIT;

  officeApps[1] := 'excel.exe';
  officeApps[2] := 'MSACESS.EXE';
  //OfficeApps[3]:= 'Winword.exe'; 
  //OfficeApps[4]:= 'powerpnt.exe'; 
  //OfficeApps[5]:= 'OUTLOOK.exe';

  for i := 1 to High(officeApps) do
  begin
    keyFound := RegQueryStringValue(HKLM, APPPATHSKEY + '\' + officeApps[i], '', officeAppPath);
    if keyFound then Break;
  end;

  if keyFound then
  begin
    // find the bitness of the application binary
    if GetBinaryType(officeAppPath, binaryType) then
      Result := binaryType;
  end;
end;

function InstallX64(): Boolean;
begin
  Result := (GetOfficeBitness = OFFICE_64_BIT);
end;

function InstallX32: Boolean;
begin
  Result := IsWin64 and (GetOfficeBitness = OFFICE_32_BIT);
end;

// The following three functions from: 
// https://github.com/florentbr/SeleniumBasic/blob/master/SeleniumBasicSetup.iss 
// Computer\HKEY_CLASSES_ROOT\Excel.Application\CurVer 
// Computer\HKEY_CLASSES_ROOT\Access.Application\CurVer

function HasExcel(): Boolean;
begin
  Result := RegKeyExists(HKCR, 'Excel.Application');
end;

function HasAccess(): Boolean;
begin
  Result := RegKeyExists(HKCR, 'Access.Application');
end;

function BoolToStr(const value: Boolean): string;
begin
  if value then
    Result := 'True'
  else
    Result := 'False';
end;

// Office version detection notes 
// Only if 2007 then we can rule out 64 bit. 2010, 2013, 2016, 
// and 365 all have both 32 and 64 bit versions 
// MS not supporting 2013 after April 2023 
// first get the path to executables: 
// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe 
// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE 
// these above should yield something like: 
// C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE 
// C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE 
// then use GetBinaryType to discover bitness 
// Check if the OS/Office requirements have been met - if not warn user

function InitializeSetup(): Boolean;
var
  officeBitness: Integer;
  answer: Integer;
begin
  if not IsWin64 then
  begin
    answer := MsgBox(
      'Setup has determined that your OS is not 64-bit Windows, which is a requirement of this installation. Do you still want to proceed?',
      mbConfirmation, MB_YESNO);
    if answer = IDYES then
      Result := True
    else
    begin
      Result := False;
      Exit;
    end;
  end;

  officeBitness := GetOfficeBitness;

  case officeBitness of
    OFFICE_UNKNOWN_BIT:
    begin
      answer := MsgBox(
        'MS Office bitness could not be determined. Are you sure that you want to proceed with the installation?',
        mbConfirmation, MB_YESNO);
      if answer = IDYES then
        Result := True
      else
        Result := False;
    end;
    OFFICE_32_BIT:
      Result := True;
    OFFICE_64_BIT:
      Result := True;
  else
    begin
      answer := MsgBox(
        'The installed version of MS Office was found but is not compatible with this installation. Are you sure that you want to proceed?',
        mbConfirmation, MB_YESNO);
      if answer = IDYES then
        Result := True
      else
        Result := False;
    end;
  end;
end;

function GetOfficeVersion(app: string): string;
var
  ver: string;
  i: Integer;
begin
  if RegQueryStringValue(HKCR, app + '.Application\CurVer', '', ver) then
  begin
    for i := 1 to Length(ver) do
    begin
      if (ver[i] >= '0') and (ver[i] <= '9') then
        Result := Result + ver[i];
    end;
  end;
end;

// The following workflow is based on Bill Stewart's UninsIS DLL version
// see https://github.com/Bill-Stewart/UninsIS 
// Refactored/simplified to run directly as IS script (versus pre-compiled DLL) 

// Get Registry path root
function GetUninstallRegRoot(isAdmin: Boolean): Integer;
begin
  if isAdmin then
    Result := HKEY_LOCAL_MACHINE
  else
    Result := HKEY_CURRENT_USER;
end;

// Builds IS uninstall Registry key
function GetUninstallRegKey(appId: string; is64Bit, isAdmin: Boolean): string;
begin
  if is64Bit and IsWin64 then
    Result := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + appId + '_is1'
  else
    Result := 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\' + appId + '_is1';

  if not IsWin64 then
    Result := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + appId + '_is1';
end;

// Determines if IS uninstall Registry key exists
function IsISPackageInstalled(appId: string; is64Bit, isAdmin: Boolean): Boolean;
begin
  Result := RegKeyExists(GetUninstallRegRoot(isAdmin), GetUninstallRegKey(appId, is64Bit, isAdmin));
end;

// Reads the IS uninstall Registry key value holding the application install version
function GetISPackageVersion(appId: string; is64Bit, isAdmin: Boolean): string;
begin
  if not RegQueryStringValue(GetUninstallRegRoot(isAdmin),
    GetUninstallRegKey(appId, is64Bit, isAdmin), 'DisplayVersion', Result) then
    Result := '';
end;

// Extracts the Nth numeric component from a dotted version string
function ParseVersionPart(const ver: string; index: Integer): Integer;
var
  p, startPos, partIndex: Integer;
  s: string;
begin
  Result := 0;
  startPos := 1;
  partIndex := 0;

  for p := 1 to Length(ver) + 1 do
  begin
    if (p > Length(ver)) or (ver[p] = '.') then
    begin
      if partIndex = index then
      begin
        s := Copy(ver, startPos, p - startPos);
        Result := StrToIntDef(s, 0);
        Exit;
      end;
      Inc(partIndex);
      startPos := p + 1;
    end;
  end;
end;

// Compare two dotted version strings
function CompareVersionStrings(const ver1, ver2: string): Integer;
var
  i, v1, v2: Integer;
begin
  Result := 0;
  for i := 0 to 3 do
  begin
    v1 := ParseVersionPart(ver1, i);
    v2 := ParseVersionPart(ver2, i);
    if v1 < v2 then
    begin
      Result := -1;
      Exit;
    end
    else if v1 > v2 then
    begin
      Result := 1;
      Exit;
    end;
  end;
end;

// Returns: 
// < 0 if version we are installing is < installed version 
// 0 if version we are installing is = installed version 
// > 0 if version we are installing is > installed version
function CompareISPackageVersion(appId, installingVersion: string; is64Bit, isAdmin: Boolean): Integer;
var
  installedVersion: string;
begin
  installedVersion := GetISPackageVersion(appId, is64Bit, isAdmin);
  if installedVersion = '' then
    Result := 0
  else
    Result := CompareVersionStrings(installingVersion, installedVersion);
end;

// Returns the path of the current (pre-existing) unins000.exe file
function GetUninstallString(appId: string; is64Bit, isAdmin: Boolean): string;
begin
  if not RegQueryStringValue(GetUninstallRegRoot(isAdmin),
    GetUninstallRegKey(appId, is64Bit, isAdmin), 'UninstallString', Result) then
    Result := '';
end;

// Returns true if package is detected as uninstalled, or false otherwise
function UninstallISPackage(appId: string; is64Bit, isAdmin: Boolean): Boolean;
var
  uninstExe: string;
  resultCode, i: Integer;
  uninstallArgs: string;
begin
  Result := False;
  uninstExe := GetUninstallString(appId, is64Bit, isAdmin);
  if uninstExe = '' then
    Exit;

  // Extract the EXE path from the uninstall string (strip quotes if present)
  if (Length(uninstExe) > 0) and (uninstExe[1] = '"') then
  begin
    Delete(uninstExe, 1, 1);
    uninstExe := Copy(uninstExe, 1, Pos('"', uninstExe) - 1);
  end
  else
  begin
    // If no quotes, take up to first space
    if Pos(' ', uninstExe) > 0 then
      uninstExe := Copy(uninstExe, 1, Pos(' ', uninstExe) - 1);
  end;

  // Run the uninstaller silently
  uninstallArgs := '/SILENT /SUPPRESSMSGBOXES /NORESTART';
  if Exec(uninstExe, uninstallArgs, '', SW_SHOW, ewWaitUntilTerminated, resultCode) then
  begin
    if resultCode = 0 then
    begin
      // Wait for the uninstaller EXE to delete itself (max ~30 seconds)
      for i := 0 to 300 do
      begin
        if not FileExists(uninstExe) then
        begin
          Result := True;
          Break;
        end;
        Log('Waiting for Uninstall to complete');
        Sleep(100);
      end;
    end;
  end;
end;

function PrepareToInstall(var NeedsRestart: Boolean): string;
var
  oldInstallDir: string;
  newInstallDir: string;
  isDifferentLocation: Boolean;
  isDifferentVersion: Boolean;
  isInstalled: Boolean;
  alwaysUninstall: Boolean;
  res: Boolean;
  is64BitInstall: Boolean;
  isAdminInstall: Boolean;
begin
  // Set alwaysUninstall to true if uninstall
  // should run regardless of version and location compare

  // if IsISPackageInstalled() then 
  // begin 
  // Version := GetISPackageVersion(); 
  // MsgBox('Package installed; version = ' + Version, mbInformation, MB_OK); 
  // end;

  alwaysUninstall := False;
  isDifferentLocation := False;
  isDifferentVersion := False;

  oldInstallDir := WizardForm.PrevAppDir;
  newInstallDir := ExpandConstant('{app}');
  if oldInstallDir <> newInstallDir then
    isDifferentLocation := True;

  is64BitInstall := Is64BitInstallMode();
  isAdminInstall := IsAdminInstallMode();

  if CompareISPackageVersion('{#AppGUID}', '{#AppVersion}', is64BitInstall, isAdminInstall) <> 0 then
    isDifferentVersion := True;

  isInstalled := IsISPackageInstalled('{#AppGUID}', is64BitInstall, isAdminInstall);

  if isInstalled then
  begin
    if isDifferentLocation or isDifferentVersion or alwaysUninstall then
    begin
      res := UninstallISPackage('{#AppGUID}', is64BitInstall, isAdminInstall);
      if res then
        Log('Previous version successfully uninstalled')
      else
        Log('Previous version not uninstalled');
    end;
  end;

  // Log some results for debugging
  Log('Old Install Dir= ' + oldInstallDir);
  Log('New Install Dir= ' + newInstallDir);
  Log('Is Admin install mode= ' + BoolToStr(IsAdminInstallMode()));
  Log('Is 64-bit install mode= ' + BoolToStr(Is64BitInstallMode()));
  Log('Is DLL already installed= ' + BoolToStr(isInstalled));
  Log('Is DLL install location different than previous= ' + BoolToStr(isDifferentLocation));
  Log('Is DLL version different than previous= ' + BoolToStr(isDifferentVersion));

  Result := '';
end;







