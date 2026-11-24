// This installer allows user to (re)install an ActiveX DLL in a location
// of their choice.
//
// It always uninstalls any previous version of the DLL before
// installing a new one. This ensures clean COM/ActiveX registration,
// handles bitness changes correctly (32-bit vs 64-bit Office), and produces
// a consistent file set regardless of which components were previously installed.
//
// Both the 32-bit and 64-bit builds of the DLL are installed and registered
// side by side. They share the same CLSID/ProgID/TypeLib GUIDs but register
// into different registry views (native 64-bit view vs. the WOW6432Node
// 32-bit view), so they never conflict. This lets the component be consumed
// by 64-bit Office, 32-bit Office, and dual-bitness COM hosts such as
// twinBASIC in either 32-bit or 64-bit mode. Each consumer automatically
// binds to the build matching its own bitness.
//
// If installing for first time, or installing after a manual uninstall, then user is prompted
// for installation type (ALLUSERS vs CURRENTUSER). If reinstalled over a previous install, 
// then the previous install type is assumed. If user wants to change type of install by 
// clicking on [installer name].exe setup file then user must manually uninstall first to
// get the prompt for installation type.
//
// The installer can be run from command line:
// [installer name].exe /CURRENTUSER
// [installer name].exe /ALLUSERS
//
// An install log can be produced using the following from command line:
// [installer name].exe /LOG="install_log.txt"
//
// Office application support is configured via the flags in the #define block
// at the top of this script (below). OFFICE_AWARE is the master switch: set it
// to "False" for a non-Office DLL and all Office-specific behavior is disabled
// -- the "supported app" detection prompt is skipped and no "Trusted Locations"
// registry entries are written, regardless of the other flags. When
// OFFICE_AWARE is "True", the SUPPORT_* flags select which Office apps the DLL
// targets (affecting the detection message and which Trusted Location entries
// are written), and REGISTER_TRUSTED_LOCS toggles those entries on or off.
// Note that both DLL bitnesses are installed regardless of which Office bitness
// (if any) is detected.
//
// ---------------------------------------------------------------------------
// REUSING THIS SCRIPT AS A TEMPLATE
// This script works for both Office-targeted and general-purpose ActiveX DLLs.
// To repurpose it for a different DLL:
//   1. Update the #define block below (AppName, AppGUID, URLs, DLL paths, etc.).
//      *** AppGUID MUST be unique per product. *** It is used as Inno's AppId,
//      which keys the uninstall-tracking registry entry and the
//      "uninstall previous version" logic in PrepareToInstall. Two DLLs sharing
//      an AppGUID would treat each other as a "previous version" and uninstall
//      one another. Generate a fresh GUID for every component.
//   2. Set the behavior flags in the #define block at the top of the script:
//        Office DLL w/ example docs : OFFICE_AWARE="True",  REGISTER_TRUSTED_LOCS="True"
//        Office DLL, no example docs: OFFICE_AWARE="True",  REGISTER_TRUSTED_LOCS="False"
//        General / non-Office DLL   : OFFICE_AWARE="False"  (SUPPORT_* and
//                                     REGISTER_TRUSTED_LOCS are then ignored)
//   3. The [Files] section expects both a _win32 and a _win64 build. For a
//      genuinely single-bitness component, comment out the Source line for the
//      bitness you don't ship; everything else still works.
// The Office-specific helpers and the [Registry] trusted-location entries can be
// left in place even for a general DLL -- with OFFICE_AWARE="False" nothing
// Office-related runs and nothing is written, so they are inert.
// ---------------------------------------------------------------------------

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
#define INIFilePath "..\Build\SeleniumVBA.ini"
#define UtilitiesPath "..\..\Utilities"
#define LogoFilePath ".\logo_setup.bmp"
#define RequirementsFilePath ".\readme.txt"
#define SetupOutputFolderPath "..\..\..\dist" 
#define AppVersion GetVersionNumbersString(DLL64FilePath)

// ---- Office app support configuration (user settings) ----
// Master switch for Office-specific behavior (e.g. the Office-detection prompt).
//   Office DLL shipping example docs : OFFICE_AWARE="True",  REGISTER_TRUSTED_LOCS="True"
//   Office DLL, no example docs      : OFFICE_AWARE="True",  REGISTER_TRUSTED_LOCS="False"
//   General / non-Office DLL         : OFFICE_AWARE="False", REGISTER_TRUSTED_LOCS="False"
#define OFFICE_AWARE "True"

// Only applicable if OFFICE_AWARE is "True"
// Set each to "True" or "False" (KEEP THE QUOTES). They affect the user-facing
// "supported apps" message and which Trusted Location registry entries are
// written. Both DLL bitnesses are always installed regardless of these flags.
#define SUPPORT_EXCEL   "True"
#define SUPPORT_ACCESS  "True"
#define SUPPORT_WORD    "False"
#define SUPPORT_PPT     "False"
#define SUPPORT_OUTLOOK "False"

// Only applicable if OFFICE_AWARE is "True"
// When "True", adds {app}\examples as a trusted location for each supported and
// installed Office app.
#define REGISTER_TRUSTED_LOCS "True"

[Setup]
AppId={{#AppGUID}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}

; Install mode (per-user vs per-machine):
;   PrivilegesRequired=lowest -> default to per-user install (no UAC at launch)
;   PrivilegesRequiredOverridesAllowed=dialog commandline -> user can override via:
;     - the "Select Setup Install Mode" dialog at first launch (no prior install detected), OR
;     - command-line switches /CURRENTUSER or /ALLUSERS
;   On subsequent runs over a prior install, the existing install's mode is inferred
;   and the dialog is suppressed. To switch modes, uninstall first via Settings > Apps.
;
;   NOTE on COM registration and privileges: per-machine (admin) installs register
;   the DLLs into HKLM\Software\Classes (visible to all users). Per-user installs run
;   non-elevated and rely on the DLL's DllRegisterServer falling back to
;   HKCU\Software\Classes. twinBASIC-built ActiveX DLLs generally support per-user
;   (HKCU) self-registration, so this normally works; if a per-user install does not
;   make the component visible to a COM host, re-run as administrator (Install for all
;   users). This applies equally to both bitnesses.
;
; DefaultDirName uses {autopf} which auto-resolves to:
;   {commonpf} = "C:\Program Files" for per-machine install, or
;   {userpf}   = "C:\Users\<user>\AppData\Local\Programs" for per-user install.

PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog commandline
DefaultDirName={autopf}\{#AppName}

DefaultGroupName={#AppName}
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
; The installer still runs only on 64-bit Windows and installs in 64-bit mode.
; This is required so the 64-bit DLL can be registered natively. The 32-bit DLL
; is registered via the "32bit" file flag, which directs Inno to use the 32-bit
; registration server (writing into the WOW6432Node view) even though the overall
; install runs in 64-bit mode.
ArchitecturesAllowed=x64os
ArchitecturesInstallIn64BitMode=x64os
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
Name: "pkg_core"; Description: "{#AppName} ActiveX Dll"; Types: full compact custom; Flags: fixed;
Name: "pkg_utils";  Description: "PowerShell Utilities"; Types: full compact custom; Flags: fixed;
Name: "pkg_docs";  Description: "MS Office Test Documents"; Types: full compact custom;
  
[Messages]
// WelcomeLabel2=This will install [name/ver] on your computer.%n%nIt is recommended that you close all other applications before continuing.
FinishedLabel=Setup has finished installing [name] on your computer. A shortcut to the DLL folder can be found on your Desktop.
ClickFinish=Click Finish to complete and exit Setup.

[Files]
; Both bitnesses are always installed and registered. The "64bit" and "32bit"
; flags are REQUIRED here: because the install runs in 64-bit mode
; (ArchitecturesInstallIn64BitMode=x64os), Inno's default registration mode would
; otherwise be 64-bit for BOTH files, which would fail to register the 32-bit
; DLL (a 32-bit DLL cannot be loaded into the 64-bit registration host). The
; explicit flags make Inno use the matching 32-/64-bit registration server, so
; each build registers into its own registry view with no conflict.
Source: {#DLL64FilePath}; DestDir: {app};  Flags: ignoreversion regserver 64bit; Components: pkg_core;
Source: {#DLL32FilePath}; DestDir: {app};  Flags: ignoreversion regserver 32bit; Components: pkg_core;
Source: {#INIFilePath}; DestDir: {app};  Flags: ignoreversion uninsneveruninstall onlyifdoesntexist; Check: IsWin64; Components: pkg_core;
Source: {#TestFolderPath}\SeleniumVBA test subs for tB.xlsm; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs; 
Source: {#TestFolderPath}\SeleniumVBA test subs for tB.accdb; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#TestFolderPath}\readme.txt; DestDir: {app}\examples; Flags: ignoreversion; Components: pkg_docs;
Source: {#UtilitiesPath}\cleanup_drivers.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\create_update_ini_file.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\launch_chrome_in_debugger_mode.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\launch_edge_in_debugger_mode.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#UtilitiesPath}\analize_registry.ps1; DestDir: {app}\utilities; Flags: ignoreversion; Components: pkg_utils;
Source: {#LicenseFilePath} ; DestDir: "{app}"; Flags: ignoreversion; Components: pkg_core;
Source: {#RequirementsFilePath} ; DestDir: "{app}"; Flags: ignoreversion; Components: pkg_core;

[Icons]
Name: "{autodesktop}\{#AppName} - Shortcut"; Filename: "{app}"
Name: "{app}\wiki help documentation"; Filename: "{#AppHelpURL}"

[Run]

[Registry]
; Add Office app trusted locations for sample files.
; These are intentionally HKCU even on per-machine installs because
; Office reads trusted locations from the current user's hive.
; Each app has three values per trusted location: Path, AllowSubFolders, Description.
; Only the Path entry needs uninsdeletekey — that flag removes the whole subkey
; on uninstall, taking the other values with it.
; Entries for each app are gated at install time by the corresponding Has* function,
; which combines the SUPPORT_* compile-time flag with a runtime check that the app
; is actually installed on the system.

; Excel
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Excel}"; \
  ValueName: "Path"; ValueType: String; ValueData: "{app}\examples"; \
  Flags: uninsdeletekey; Check: HasExcel and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Excel}"; \
  ValueName: "AllowSubFolders"; ValueType: DWord; ValueData: "1"; \
  Check: HasExcel and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Excel}"; \
  ValueName: "Description"; ValueType: String; ValueData: "{#AppName} example documents"; \
  Check: HasExcel and ShouldRegisterTrustedLocs;

; Access
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Access}"; \
  ValueName: "Path"; ValueType: String; ValueData: "{app}\examples"; \
  Flags: uninsdeletekey; Check: HasAccess and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Access}"; \
  ValueName: "AllowSubFolders"; ValueType: DWord; ValueData: "1"; \
  Check: HasAccess and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Access}"; \
  ValueName: "Description"; ValueType: String; ValueData: "{#AppName} example documents"; \
  Check: HasAccess and ShouldRegisterTrustedLocs;

; Word
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Word}"; \
  ValueName: "Path"; ValueType: String; ValueData: "{app}\examples"; \
  Flags: uninsdeletekey; Check: HasWord and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Word}"; \
  ValueName: "AllowSubFolders"; ValueType: DWord; ValueData: "1"; \
  Check: HasWord and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|Word}"; \
  ValueName: "Description"; ValueType: String; ValueData: "{#AppName} example documents"; \
  Check: HasWord and ShouldRegisterTrustedLocs;

; PowerPoint
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|PowerPoint}"; \
  ValueName: "Path"; ValueType: String; ValueData: "{app}\examples"; \
  Flags: uninsdeletekey; Check: HasPPT and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|PowerPoint}"; \
  ValueName: "AllowSubFolders"; ValueType: DWord; ValueData: "1"; \
  Check: HasPPT and ShouldRegisterTrustedLocs;
Root: HKCU; Subkey: "{code:GetTrustedLocSubkey|PowerPoint}"; \
  ValueName: "Description"; ValueType: String; ValueData: "{#AppName} example documents"; \
  Check: HasPPT and ShouldRegisterTrustedLocs;

; Outlook
  // Note: Outlook does not honor "Trusted Locations" the way the document-based
  // Office apps do — it uses different macro-security mechanisms. The SUPPORT_OUTLOOK
  // flag affects the supported-apps message, but no trusted-location registry
  // entries are written for Outlook even when this is True.

[Code]
const
  // ---- Office app support configuration ----
  // These values are configured via the #define block at the TOP of the script
  // (SUPPORT_EXCEL, SUPPORT_ACCESS, ...). The lines below just inherit those
  // compile-time defines; edit the flags up top, not here.
  // Only relevant when OFFICE_AWARE is "True". They select which Office apps
  // the DLL targets, affecting the user-facing "supported apps" message in
  // InitializeSetup and which Trusted Location registry entries are written.
  // They do not control which DLL bitness is installed -- both bitnesses are
  // always installed.
  SUPPORT_EXCEL   = {#SUPPORT_EXCEL};
  SUPPORT_ACCESS  = {#SUPPORT_ACCESS};
  SUPPORT_WORD    = {#SUPPORT_WORD};
  SUPPORT_PPT     = {#SUPPORT_PPT};
  SUPPORT_OUTLOOK = {#SUPPORT_OUTLOOK};

  // ---- Office-aware behavior ----
  // Configured via the #define block at the TOP of the script (OFFICE_AWARE).
  // Master switch for Office-specific behavior. When "False" (for DLLs that
  // have nothing to do with Office): the Office-detection prompt in
  // InitializeSetup is skipped, AND no trusted-location entries are written --
  // ShouldRegisterTrustedLocs gates on this flag, so SUPPORT_* and
  // REGISTER_TRUSTED_LOCS are effectively ignored. The install/registration of
  // both DLL bitnesses is unaffected either way.
  OFFICE_AWARE = {#OFFICE_AWARE};

  // ---- Trusted Locations registration ----
  // Configured via the #define block at the TOP of the script
  // (REGISTER_TRUSTED_LOCS). Only takes effect when OFFICE_AWARE is "True".
  // When both are "True", the installer adds {app}\examples as a trusted
  // location for each supported and installed Office app. Set to "False" for
  // DLLs that don't ship example documents or don't need their install folder
  // trusted. (The Has* runtime checks also gate these, so nothing is written
  // for an app that isn't installed even when this is "True".)
  REGISTER_TRUSTED_LOCS = {#REGISTER_TRUSTED_LOCS};

  // ---- Other constants ----
  APPPATHSKEY = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths';
  OFFICE_UNKNOWN_BIT = -1;
  OFFICE_32_BIT = 0;
  OFFICE_64_BIT = 6;
  UNINSTALL_ARGS = '/SILENT /SUPPRESSMSGBOXES /NORESTART';
  MAX_WAIT_ITERATIONS = 300;  // 300 * 100ms = 30 seconds
  WAIT_INTERVAL_MS = 100;
  UNINSTALL_NATIVE_PATH = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\';
  UNINSTALL_WOW64_PATH  = 'Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\';

// API call to determine bitness of an executable
function GetBinaryType(ApplicationName: string; var BinaryType: Integer): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

// Appends a string to a dynamic string array.
procedure AppendToArray(var arr: TArrayOfString; const value: string);
begin
  SetArrayLength(arr, GetArrayLength(arr) + 1);
  arr[GetArrayLength(arr) - 1] := value;
end;

// Returns the value of a SUPPORT_* constant. Wrapping the constant in a
// function defeats Pascal Script's compile-time constant folding, which
// would otherwise produce "True and ..." or "False and ..." warnings on
// every Has* function depending on the configured support flags.
function IsSupported(flag: Boolean): Boolean;
begin
  Result := flag;
end;

// Returns True if trusted-location registry entries should be written for
// supported Office apps. Gated on OFFICE_AWARE so that turning off Office
// behavior at the top also suppresses the trusted-location writes, regardless
// of REGISTER_TRUSTED_LOCS. Wrapped in a function (rather than referenced as a
// constant directly) to avoid Pascal Script's constant-folding warnings.
function ShouldRegisterTrustedLocs(): Boolean;
begin
  Result := IsSupported(OFFICE_AWARE) and IsSupported(REGISTER_TRUSTED_LOCS);
end;

// Returns a human-readable, comma-separated list of supported Office apps,
// based on the SUPPORT_* constants. Used in user-facing diagnostic messages.
function GetSupportedAppsList(): string;
begin
  Result := '';
  if IsSupported(SUPPORT_EXCEL)   then Result := Result + ', Excel';
  if IsSupported(SUPPORT_ACCESS)  then Result := Result + ', Access';
  if IsSupported(SUPPORT_WORD)    then Result := Result + ', Word';
  if IsSupported(SUPPORT_PPT)     then Result := Result + ', PowerPoint';
  if IsSupported(SUPPORT_OUTLOOK) then Result := Result + ', Outlook';
  // Strip the leading ", " if anything was appended
  if Length(Result) > 0 then
    Result := Copy(Result, 3, Length(Result));
end;

// Office version detection notes: 
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
//
// Bitness is now used only for the informational InitializeSetup message — it
// no longer determines which DLL is installed, since both bitnesses are always
// installed and registered.
function GetOfficeBitness(): Integer;
var
  officeApps: TArrayOfString;
  i: Integer;
  officeAppPath: string;
  binaryType: Integer;
  keyFound: Boolean;
begin
  Result := OFFICE_UNKNOWN_BIT;

  // Build the list of Office app executables to probe, in priority order.
  // Only apps marked as supported by the SUPPORT_* flags are checked.
  if IsSupported(SUPPORT_EXCEL)   then AppendToArray(officeApps, 'excel.exe');
  if IsSupported(SUPPORT_ACCESS)  then AppendToArray(officeApps, 'MSACCESS.EXE');
  if IsSupported(SUPPORT_WORD)    then AppendToArray(officeApps, 'winword.exe');
  if IsSupported(SUPPORT_PPT)     then AppendToArray(officeApps, 'powerpnt.exe');
  if IsSupported(SUPPORT_OUTLOOK) then AppendToArray(officeApps, 'outlook.exe');

  keyFound := False;
  for i := 0 to GetArrayLength(officeApps) - 1 do
  begin
    keyFound := RegQueryStringValue(HKLM, APPPATHSKEY + '\' + officeApps[i], '', officeAppPath);
    if keyFound then Break;
  end;

  if keyFound then
  begin
    // Found a supported Office app — determine its bitness via the binary header
    if GetBinaryType(officeAppPath, binaryType) then
      Result := binaryType;
  end;
end;

// Has* functions: each combines the compile-time SUPPORT_* flag with a runtime
// check that the corresponding Office app is actually installed on the system.
// Used as Check: gates on [Registry] entries so that unsupported or absent apps
// don't get trusted-location entries written.
function HasExcel(): Boolean;
begin
  Result := IsSupported(SUPPORT_EXCEL) and RegKeyExists(HKCR, 'Excel.Application');
end;

function HasAccess(): Boolean;
begin
  Result := IsSupported(SUPPORT_ACCESS) and RegKeyExists(HKCR, 'Access.Application');
end;

function HasWord(): Boolean;
begin
  Result := IsSupported(SUPPORT_WORD) and RegKeyExists(HKCR, 'Word.Application');
end;

function HasPPT(): Boolean;
begin
  Result := IsSupported(SUPPORT_PPT) and RegKeyExists(HKCR, 'PowerPoint.Application');
end;

function HasOutlook(): Boolean;
begin
  Result := IsSupported(SUPPORT_OUTLOOK) and RegKeyExists(HKCR, 'Outlook.Application');
end;

function BoolToStr(const value: Boolean): string;
begin
  if value then
    Result := 'True'
  else
    Result := 'False';
end;

function InitializeSetup(): Boolean;
var
  officeBitness: Integer;
  answer: Integer;
  supportedList: string;
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
  supportedList := GetSupportedAppsList();

  // Non-Office DLLs skip all Office detection and messaging — a machine with
  // no Office is a perfectly valid target (e.g. developing against the
  // component in twinBASIC), so don't prompt about it.
  if not IsSupported(OFFICE_AWARE) then
  begin
    Result := True;
    Exit;
  end;

  // Both 32-bit and 64-bit DLLs are installed and registered regardless of the
  // detected Office bitness, so the prompts below are purely informational. A
  // user with no Office installed (e.g. developing against the component in
  // twinBASIC) can proceed normally.
  case officeBitness of
    OFFICE_UNKNOWN_BIT:
    begin
      answer := MsgBox(
        'None of the supported MS Office applications could be detected on this system, ' +
        'or their bitness could not be determined.' + #13#10 + #13#10 +
        'This installation supports: ' + supportedList + '.' + #13#10 + #13#10 +
        'Both the 32-bit and 64-bit versions of the DLL will still be installed and ' +
        'registered, so the component remains usable from other COM hosts such as ' +
        'twinBASIC (in either 32-bit or 64-bit mode).' + #13#10 + #13#10 +
        'Do you want to proceed with the installation?',
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
        'The installed version of MS Office was found but its bitness is not one of the ' +
        'expected values. Both DLL bitnesses will be installed regardless. ' +
        'Do you want to proceed?',
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

function GetTrustedLocSubkey(app: string): string;
var
  version: string;
begin
  version := GetOfficeVersion(app);
  if version = '' then
    Result := ''  // app not installed; caller's Check: will skip this entry anyway
  else
    Result :=
      'Software\Microsoft\Office\' + version + '.0\' +
      app + '\Security\Trusted Locations\' +
      '{#AppName}';
end;
 
// Get Registry path root
function GetUninstallRegRoot(isAdmin: Boolean): Integer;
begin
  if isAdmin then
    Result := HKEY_LOCAL_MACHINE
  else
    Result := HKEY_CURRENT_USER;
end;

// Builds the Inno Setup uninstall registry key path for a given AppID.
// Inno writes its uninstall info under different paths depending on the
// installation mode (per-user vs per-machine) and the installer's bitness.
//
//   HKCU (per-user):   always Software\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}_is1
//                      (HKCU has no WOW64 redirection)
//   HKLM (per-machine):
//     64-bit installer on 64-bit Windows -> Software\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}_is1
//     32-bit installer on 64-bit Windows -> Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}_is1
//     32-bit Windows                     -> Software\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}_is1
function GetUninstallRegKey(appId: string; is64Bit, isAdmin: Boolean): string;
begin
  // HKCU has no WOW64 redirection — same path regardless of bitness
  if not isAdmin then
    Result := UNINSTALL_NATIVE_PATH + appId + '_is1'
  // HKLM on 64-bit Windows: bitness of installer determines the view
  else if IsWin64 and (not is64Bit) then
    Result := UNINSTALL_WOW64_PATH + appId + '_is1'
  // HKLM on 64-bit Windows running 64-bit installer, or HKLM on 32-bit Windows
  else
    Result := UNINSTALL_NATIVE_PATH + appId + '_is1';
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

// Compares the version being installed against a given installed version string.
// Returns:
//   < 0 if version we are installing is < installed version
//   = 0 if version we are installing is = installed version, or if installedVersion
//         is empty (no prior install detected)
//   > 0 if version we are installing is > installed version
function CompareVersionToInstalled(installingVersion, installedVersion: string): Integer;
var
  installingPacked, installedPacked: Int64;
begin
  if installedVersion = '' then
  begin
    Result := 0;
    Exit;
  end;

  if not StrToVersion(installingVersion, installingPacked) then installingPacked := 0;
  if not StrToVersion(installedVersion, installedPacked) then installedPacked := 0;
  Result := ComparePackedVersion(installingPacked, installedPacked);
end;

// Returns the path of the current (pre-existing) unins000.exe file
function GetUninstallString(appId: string; is64Bit, isAdmin: Boolean): string;
begin
  if not RegQueryStringValue(GetUninstallRegRoot(isAdmin),
    GetUninstallRegKey(appId, is64Bit, isAdmin), 'UninstallString', Result) then
    Result := '';
end;

// Strips surrounding double quotes from a string if present.
// "C:\Program Files\App\unins000.exe" -> C:\Program Files\App\unins000.exe
// C:\Apps\unins000.exe                 -> C:\Apps\unins000.exe (unchanged)
function StripQuotes(const s: string): string;
begin
  Result := s;
  if (Length(Result) >= 2) and (Result[1] = '"') and (Result[Length(Result)] = '"') then
    Result := Copy(Result, 2, Length(Result) - 2);
end;

// Returns true if package is detected as uninstalled, or false otherwise
function UninstallISPackage(appId: string; is64Bit, isAdmin: Boolean): Boolean;
var
  uninstExe: string;
  resultCode, i: Integer;

begin
  Result := False;

  uninstExe := GetUninstallString(appId, is64Bit, isAdmin);
  if uninstExe = '' then
  begin
    Log('UninstallISPackage: no UninstallString found in registry');
    Exit;
  end;

  uninstExe := StripQuotes(Trim(uninstExe));

  if not FileExists(uninstExe) then
  begin
    Log('UninstallISPackage: uninstaller not found at ' + uninstExe);
    Exit;
  end;

  Log('UninstallISPackage: running ' + uninstExe + ' ' + UNINSTALL_ARGS);

  if not Exec(uninstExe, UNINSTALL_ARGS, '', SW_HIDE, ewWaitUntilTerminated, resultCode) then
  begin
    Log('UninstallISPackage: Exec failed to launch uninstaller');
    Exit;
  end;

  if resultCode <> 0 then
  begin
    Log('UninstallISPackage: uninstaller returned non-zero exit code ' + IntToStr(resultCode));
    Exit;
  end;

  // Inno's uninstaller spawns a helper that deletes unins000.exe after the
  // main process exits, so we poll for the file to disappear.
  for i := 1 to MAX_WAIT_ITERATIONS do
  begin
    if not FileExists(uninstExe) then
    begin
      Log('UninstallISPackage: uninstall confirmed after ' + IntToStr(i * WAIT_INTERVAL_MS) + 'ms');
      Result := True;
      Exit;
    end;
    Sleep(WAIT_INTERVAL_MS);
  end;

  Log('UninstallISPackage: timed out waiting for uninstaller to delete itself');
end;

function PrepareToInstall(var NeedsRestart: Boolean): string;
var
  oldInstallDir: string;
  newInstallDir: string;
  isInstalledPerUser: Boolean;
  isInstalledPerMachine: Boolean;
  isInstalled: Boolean;
  res: Boolean;
  is64BitInstall: Boolean;
  newIsAdminInstall: Boolean;
  uninstallAsAdmin: Boolean;
  versionCompare: Integer;
  installedVersion: string;
  answer: Integer;
begin
  oldInstallDir := WizardForm.PrevAppDir;
  newInstallDir := ExpandConstant('{app}');

  is64BitInstall := Is64BitInstallMode();
  newIsAdminInstall := IsAdminInstallMode();

  // Check both hives — a previous install might have used a different mode
  // than the user is choosing now, and we need to clean it up either way.
  isInstalledPerUser    := IsISPackageInstalled('{#AppGUID}', is64BitInstall, False);
  isInstalledPerMachine := IsISPackageInstalled('{#AppGUID}', is64BitInstall, True);
  isInstalled := isInstalledPerUser or isInstalledPerMachine;

  // Determine which hive holds the existing install (preferring per-machine
  // if somehow both exist).
  if isInstalledPerMachine then
    uninstallAsAdmin := True
  else
    uninstallAsAdmin := False;

  // A per-user install cannot uninstall a per-machine install without
  // elevation. Abort early so that UninstallISPackage doesn't fail mid-way through.
  if isInstalledPerMachine and (not newIsAdminInstall) then
  begin
    Log('Per-machine install detected but new install is per-user — cannot proceed');
    Result := 'A previous installation of {#AppName} was made for all users (admin install). ' +
             'You must either re-run this installer as an administrator (choose "Install for all users") ' +
             'to upgrade in place, or manually uninstall the previous version first.';
    Exit;
  end;

  if isInstalled then
  begin
    // Compare versions for diagnostics and to detect downgrades.
    installedVersion := GetISPackageVersion('{#AppGUID}', is64BitInstall, uninstallAsAdmin);
    versionCompare := CompareVersionToInstalled('{#AppVersion}', installedVersion);

    // Warn if the user is about to install an older version over a newer one.
    // Skipped in silent installs (automation is assumed to know what it's doing).
    if versionCompare < 0 then
    begin
      Log('Downgrade detected: installing {#AppVersion} over installed ' + installedVersion);

      if not WizardSilent then
      begin
        answer := MsgBox(
          'A newer version of {#AppName} (' + installedVersion + ') is already installed.' + #13#10 + #13#10 +
          'You are about to install version {#AppVersion}, which is older.' + #13#10 + #13#10 +
          'The newer version will be uninstalled first. Do you want to continue?',
          mbConfirmation, MB_YESNO);
        if answer <> IDYES then
        begin
          Log('User cancelled downgrade');
          Result := 'Installation cancelled.';
          Exit;
        end;
      end;
    end;

    // Always uninstall the previous version before installing the new one.
    // This ensures clean COM registration, correct bitness, and a consistent
    // file set regardless of component selection changes.
    res := UninstallISPackage('{#AppGUID}', is64BitInstall, uninstallAsAdmin);
    if res then
      Log('Previous version successfully uninstalled')
    else
    begin
      Log('Previous version uninstall failed');
      Result := 'The previous version of {#AppName} could not be uninstalled. ' +
               'Please uninstall it manually from Windows Settings > Apps, ' +
               'then run this installer again.';
      Exit;
    end;

    // If both somehow existed, also clean up the other hive
    if isInstalledPerUser and isInstalledPerMachine then
    begin
      Log('Both per-user and per-machine installs detected — cleaning up the other hive');
      res := UninstallISPackage('{#AppGUID}', is64BitInstall, not uninstallAsAdmin);
      if res then
        Log('Secondary previous version also successfully uninstalled')
      else
        Log('Secondary previous version uninstall failed (continuing anyway)');
    end;
  end;

  // Log some results for debugging
  Log('Old Install Dir= ' + oldInstallDir);
  Log('New Install Dir= ' + newInstallDir);
  Log('New install is admin mode= ' + BoolToStr(newIsAdminInstall));
  Log('Is 64-bit install mode= ' + BoolToStr(is64BitInstall));
  Log('Was DLL already installed (per-user)= ' + BoolToStr(isInstalledPerUser));
  Log('Was DLL already installed (per-machine)= ' + BoolToStr(isInstalledPerMachine));
  if isInstalled then
    Log('Installed version was ' + installedVersion + ', new version is {#AppVersion} (compare=' + IntToStr(versionCompare) + ')');

  Result := '';
end;
