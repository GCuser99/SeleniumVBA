// This installer script references Bill Stewart's UninsIS DLL version 1.5.0.
// see https://github.com/Bill-Stewart/UninsIS
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
// The following definition points to the path of 
// Bill Stewart's UninsIS.dll
#define UninstallDLLFilePath ".\UninsIS\UninsIS.dll"

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
Source: {#LicenseFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
Source: {#RequirementsFilePath} ; DestDir: "{app}"; Flags: ignoreversion ; Components: pkg_core;
; Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme
; For importing DLL functions at setup
Source: {#UninstallDLLFilePath}; Flags: dontcopy

[Icons]
Name: "{autodesktop}\SeleniumVBA - Shortcut"; Filename: "{app}"
Name: "{app}\wiki help documentation"; Filename: "{#AppHelpURL}"

[Run]
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\utilities\create_update_ini_file.ps1"" -keepExistingValues -iniFilePath ""{app}\SeleniumVBA.ini"""; Description: "Create/update optional INI Settings file (old values will be preserved)"; Flags: postinstall skipifsilent runhidden;

[Registry]
;Add excel and access trusted location for examples
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVersion|Excel}.0\Excel\Security\Trusted Locations\SeleniumVBA"; ValueName: "Path"; ValueType: String; ValueData: "{app}\examples";  Flags: uninsdeletekey; Check: HasExcel;
Root: HKCU; Subkey: "Software\Microsoft\Office\{code:GetOfficeVersion|Access}.0\Access\Security\Trusted Locations\SeleniumVBA"; ValueName: "Path"; ValueType: String; ValueData: "{app}\examples";  Flags: uninsdeletekey; Check: HasAccess;
;Enable WScript host in case it's been disabled
;Root: HKCU; Subkey: "Software\Microsoft\Windows Script Host\Settings"; ValueName: "Enabled"; ValueType: dword; ValueData: 1; 

[Code]
const
  AppPathsKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths';
  OFFICE_UNKNOWN_BIT = -1;
  OFFICE_32_BIT = 0;
  OFFICE_64_BIT = 6;

Function GetBinaryType(ApplicationName: string; var BinaryType: Integer): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

// The following code from Martin Prikryl - see: 
// https://stackoverflow.com/questions/47431008/getting-the-version-and-platform-of-office-application-from-windows-registry#47443674
Function GetOfficeBitness():Integer;
  Var OfficeApps: array[1..2] of string; 
  Var i: Integer;
  Var OfficeAppPath: String;
  Var BinaryType: Integer;
  Var KeyFound: Boolean;
  Begin
    //find the Office app binary path from Registry - 
    //first try Excel, and if that fails then Access, then Word, etc
    //once path is found, then use API to determine bitness of app
    Result := OFFICE_UNKNOWN_BIT;

    OfficeApps[1]:= 'excel.exe';
    OfficeApps[2]:= 'MSACESS.EXE';
    //OfficeApps[3]:= 'Winword.exe';
    //OfficeApps[4]:= 'powerpnt.exe';
    //OfficeApps[5]:= 'OUTLOOK.exe';

    For i:=1 To High(OfficeApps) Do
      Begin
        KeyFound := RegQueryStringValue(HKLM, AppPathsKey + '\' + OfficeApps[i], '', OfficeAppPath);
        If KeyFound Then Break;
      End;
    
    If KeyFound Then
      Begin
        // find the bitness of the application binary 
        if GetBinaryType(OfficeAppPath, BinaryType) Then Result:=BinaryType;       
      End;
  End;

Function InstallX64(): Boolean;
  Begin
    Result := (GetOfficeBitness = OFFICE_64_BIT);
  End;

Function InstallX32: Boolean;
  Begin
    Result := IsWin64 And (GetOfficeBitness = OFFICE_32_BIT);
  End;

// The following three functions from:
// https://github.com/florentbr/SeleniumBasic/blob/master/SeleniumBasicSetup.iss
//Computer\HKEY_CLASSES_ROOT\Excel.Application\CurVer
//Computer\HKEY_CLASSES_ROOT\Access.Application\CurVer

Function HasExcel(): Boolean;
  Begin
    Result := RegKeyExists(HKCR, 'Excel.Application');
  End;

Function HasAccess(): Boolean;
  Begin
    Result := RegKeyExists(HKCR, 'Access.Application');
  End;

Function BoolToStr(const value: Boolean): String;
  Begin
    If value Then Result := 'True' Else Result := 'False';
  End;

// The following procedures are needed for UninsIS package

// Import IsISPackageInstalled() function from UninsIS.dll at setup time
function DLLIsISPackageInstalled(AppId: string; Is64BitInstallMode,
  IsAdminInstallMode: DWORD): DWORD;
  external 'IsISPackageInstalled@files:UninsIS.dll stdcall setuponly';

// Import CompareISPackageVersion() function from UninsIS.dll at setup time
function DLLCompareISPackageVersion(AppId, InstallingVersion: string;
  Is64BitInstallMode, IsAdminInstallMode: DWORD): Integer;
  external 'CompareISPackageVersion@files:UninsIS.dll stdcall setuponly';

// Import GetISPackageVersion() function from UninsIS.dll at setup time
function DLLGetISPackageVersion(AppId, Version: string;
  NumChars, Is64BitInstallMode, IsAdminInstallMode: DWORD): DWORD;
  external 'GetISPackageVersion@files:UninsIS.dll stdcall setuponly';

// Import UninstallISPackage() function from UninsIS.dll at setup time
function DLLUninstallISPackage(AppId: string; Is64BitInstallMode,
  IsAdminInstallMode: DWORD): DWORD;
  external 'UninstallISPackage@files:UninsIS.dll stdcall setuponly';

// Wrapper for UninsIS.dll IsISPackageInstalled() function
// Returns true if package is detected as installed, or false otherwise
function IsISPackageInstalled(): Boolean;
begin
  result := DLLIsISPackageInstalled('{#AppGUID}',  // AppId
    DWORD(Is64BitInstallMode()),                   // Is64BitInstallMode
    DWORD(IsAdminInstallMode())) = 1;              // IsAdminInstallMode
  if result then
    Log('UninsIS.dll - Package detected as installed')
  else
    Log('UninsIS.dll - Package not detected as installed');
end;

// Wrapper for UninsIS.dll GetISPackageVersion() function
function GetISPackageVersion(): string;
var
  NumChars: DWORD;
  OutStr: string;
begin
  result := '';
  // First call: Get number of characters needed for version string
  NumChars := DLLGetISPackageVersion('{#AppGUID}',  // AppId
    '',                                             // Version
    0,                                              // NumChars
    DWORD(Is64BitInstallMode()),                    // Is64BitInstallMode
    DWORD(IsAdminInstallMode()));                   // IsAdminInstallMode
  // Allocate string to receive output
  SetLength(OutStr, NumChars);
  // Second call: Get version number string
  if DLLGetISPackageVersion('{#AppGUID}',  // AppID
    OutStr,                                // Version
    NumChars,                              // NumChars
    DWORD(Is64BitInstallMode()),           // Is64BitInstallMode
    DWORD(IsAdminInstallMode())) > 0 then  // IsAdminInstallMode
  begin
    result := OutStr;
  end;
end;

// Wrapper for UninsIS.dll CompareISPackageVersion() function
// Returns:
// < 0 if version we are installing is < installed version
// 0   if version we are installing is = installed version
// > 0 if version we are installing is > installed version
function CompareISPackageVersion(): Integer;
begin
  result := DLLCompareISPackageVersion('{#AppGUID}',  // AppId
    '{#AppVersion}',                                  // InstallingVersion
    DWORD(Is64BitInstallMode()),                      // Is64BitInstallMode
    DWORD(IsAdminInstallMode()));                     // IsAdminInstallMode
  if result < 0 then
    Log('UninsIS.dll - This version {#AppVersion} older than installed version')
  else if result = 0 then
    Log('UninsIS.dll - This version {#AppVersion} same as installed version')
  else
    Log('UninsIS.dll - This version {#AppVersion} newer than installed version');
end;

// Wrapper for UninsIS.dll UninstallISPackage() function
// Returns 0 for success, non-zero for failure
function UninstallISPackage(): DWORD;
begin
  result := DLLUninstallISPackage('{#AppGUID}',  // AppId
    DWORD(Is64BitInstallMode()),                 // Is64BitInstallMode
    DWORD(IsAdminInstallMode()));                // IsAdminInstallMode
  if result = 0 then
    Log('UninsIS.dll - Installed package uninstall completed successfully')
  else
    Log('UninsIS.dll - installed package uninstall did not complete successfully');
end;

function PrepareToInstall(var NeedsRestart: Boolean): string;

var oldInstallDir: String; newInstallDir: String;
var isDifferentLocation: Boolean;
var isDifferentVersion: Boolean;
var isInstalled: Boolean;
var alwaysUninstall: Boolean;
begin
  // set alwaysUninstall to true if uninstall
  // should run regardless of version and location compare

  // if IsISPackageInstalled() then
  // begin
  //  Version := GetISPackageVersion();
  //  MsgBox('Package installed; version = ' + Version, mbInformation, MB_OK);
  // end;

  alwaysUninstall := false;
  isDifferentLocation := false;
  isDifferentVersion := false;

  oldInstallDir := WizardForm.PrevAppDir;
  newInstallDir := ExpandConstant('{app}');
  
  if oldInstallDir <> newInstallDir then isDifferentLocation := true;

  if CompareISPackageVersion() <> 0  then isDifferentVersion := true;

  isInstalled := IsISPackageInstalled()

  if isInstalled then
    begin
      if isDifferentLocation or isDifferentVersion or alwaysUninstall then
        UninstallISPackage();
    end;

  // log some results for debugging
  // Log('Is DLL already installed= ' + BoolToStr(isInstalled));
  // Log('Is DLL install location different than previous= ' + BoolToStr(isDifferentLocation));
  // Log('Is DLL version different than previous= ' + BoolToStr(isDifferentVersion));

  result := '';
end;

// Office version detection notes
// Only if 2007 then we can rule out 64 bit. 2010, 2013, 2016, 
// and 365 all have both 32 and 64 bit versions
// MS not supporting 2013 after April 2023

//first get the path to executables:
//HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe
//HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE
//these above should yield something like:
//C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE
//C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE
//then use GetBinaryType to discover bitness

// Check if the OS/Office requirements have been met - if not warn user 
Function InitializeSetup(): Boolean;
  Var OfficeBitness: Integer;
  Var answer: Integer;
  Begin
    if Not IsWin64 then
      Begin
        answer := MsgBox('Setup has determined that your OS is not 64-bit Windows, which is a requirement of this installation. Do you still want to proceed?', mbConfirmation, MB_YESNO); 
        If answer = IDYES Then
          Begin
            Result:=True;
          End 
        Else
          Begin
            Result := False;
            Exit;
          End; 
      End;
    OfficeBitness:= GetOfficeBitness
    Case OfficeBitness of  
      OFFICE_UNKNOWN_BIT : Begin 
        answer := MsgBox('MS Office bitness could not be determined. Are you sure that you want to proceed with the installation?', mbConfirmation, MB_YESNO); 
        if answer = IDYES then Result:=True else Result := False End;
      OFFICE_32_BIT : Begin 
        Result:=True; End; 
      OFFICE_64_BIT : Begin 
        Result:=True; End;
    Else
        Begin
          answer := MsgBox('The installed version of MS Office was found but is not compatible with this installation. Are you sure that you want to proceed?', mbConfirmation, MB_YESNO); 
          if answer = IDYES then Result:=True else Result := False;
        End;      
    End;
  End;

Function GetOfficeVersion(app: String): String;
  Var ver: String; i: Integer;
  Begin
    If RegQueryStringValue(HKCR, app + '.Application\CurVer', '', ver) Then Begin
      For i := 1 To Length(ver) Do Begin
        If (ver[i] >= '0') And (ver[i] <= '9') Then
          Result := Result + ver[i];
      End;
    End;
  End;





