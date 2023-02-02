Attribute VB_Name = "WebShared"
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v3.4
' A Selenium wrapper for Edge, Chrome, Firefox, and IE written in Windows VBA based on JSon wire protocol.
'
' (c) GCUser99
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' ==========================================================================
' For more info:
' https://docs.microsoft.com/en-us/dotnet/standard/io/file-path-formats
' http://vbnet.mvps.org/index.html?code/fileapi/pathisrelative.htm
' https://stackoverflow.com/questions/57475738/ (for use of SetCurrentDirectory)
' https://stackoverflow.com/a/72736800/11738627 (handling of OneDrive/SharePoint cloud urls)

'Several points of clarification:
'If the basePath is not specified or vbNullString, then the basePath is set to the path of active code project's
'parent document. The user has the ability to change the default basePath through DefaultIOFolder
'The only times in the code base where the default basePath is not specified is in DefaultIOFolder & DefaultDriverFolder

#Const AccessHost = False
#Const ExcelHost = True

Option Explicit
Option Private Module

'for the Sleep procedure
Private Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Public Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Public Declare PtrSafe Function getTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef counter As Currency) As Long

Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function PathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal pszPath As String) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long

Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function GetFullLocalPath(ByVal inputPath As String, Optional ByVal basePath As String = vbNullString) As String
    'Returns an absolute path from a relative path and a fully qualified base path.
    'basePath defaults to the folder path of the document that holds the Active VBA Project
    'fso.GetAbsolutePathName interprets a url as a relative path, so must avoid for url's

    Dim fso As New Scripting.FileSystemObject, savePath As String

    'make sure no rogue beginning or ending spaces and expand "%[Environ Variable]%" if in the path
    inputPath = ExpandEnvironVariable(VBA.Trim$(inputPath))

    If Not IsPathRelative(inputPath) Then 'its an absolute path
        'just in case OneDrive/SharePoint user has specified a path built with ThisWorkbook.Path...
        If IsPathHTTPS(inputPath) Then inputPath = GetLocalOneDrivePath(inputPath)
        
        'normalize the path if its not a url - this insures that path separators are correct, and
        'if a folder, has no ending separator
        If Not IsPathUrl(inputPath) Then inputPath = fso.GetAbsolutePathName(inputPath)
        
        GetFullLocalPath = inputPath
    Else 'ok then convert relative path to absolute
        'make sure no unintended beginning or ending spaces
        basePath = ExpandEnvironVariable(VBA.Trim$(basePath))
        
        If basePath = vbNullString Then
            basePath = ActiveVBAProjectFolderPath
        Else
            'it's possible that user specified a relative reference folder path - convert it to absolute relative to
            'the folder path of the document that holds the Active VBA Project
            If IsPathRelative(basePath) Then basePath = GetFullLocalPath(basePath, ActiveVBAProjectFolderPath)
        End If

        'convert OneDrive path if needed
        If IsPathHTTPS(basePath) Then basePath = GetLocalOneDrivePath(basePath)
        
        'check that reference path exists and notify user if not
        If Not fso.FolderExists(basePath) Then
            If Not IsPathHTTPS(basePath) Then 'its a url which fso doesn't support - must trust that it exists (@6DiegoDiego9)
                Err.raise 1, "WebShared", "Reference folder basePath does not exist." & vbNewLine & vbNewLine & basePath & vbNewLine & vbNewLine & "Please specify a valid folder path."
            End If
        End If
        
        'employ fso to make the conversion of relative path to absolute
        savePath = CurDir()
        SetCurrentDirectory basePath
        GetFullLocalPath = fso.GetAbsolutePathName(inputPath)
        SetCurrentDirectory savePath
    End If
End Function

Private Function GetLocalOneDrivePath(ByVal strPath As String) As String
    'thanks to @6DiegoDiego9 for doing research on this (see https://stackoverflow.com/a/72736800/11738627)
    'this function returns the original/local disk path associated with a synched OneDrive or SharePoint cloud url
    
    If IsPathHTTPS(strPath) Then
        Const HKEY_CURRENT_USER = &H80000001
        Dim objReg As WbemScripting.SWbemObjectEx 'changed to early binding by GCUser99
        Dim regPath As String
        Dim subKeys() As Variant
        Dim subKey As Variant
        Dim strValue As String
        Dim strMountpoint As String
        Dim strSecPart As String

        Static pathSep As String

        If pathSep = vbNullString Then pathSep = "\"
    
        Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

        regPath = "Software\SyncEngines\Providers\OneDrive\"
        objReg.EnumKey HKEY_CURRENT_USER, regPath, subKeys
        
        If IsArrayInitialized(subKeys) Then 'found OneDrive in registry
            For Each subKey In subKeys
                objReg.GetStringValue HKEY_CURRENT_USER, regPath & subKey, "UrlNamespace", strValue
                If InStr(strPath, strValue) > 0 Then
                    objReg.GetStringValue HKEY_CURRENT_USER, regPath & subKey, "MountPoint", strMountpoint
                    strSecPart = Replace$(Mid$(strPath, Len(strValue)), "/", pathSep)
                    GetLocalOneDrivePath = strMountpoint & strSecPart
        
                    Do Until Dir(GetLocalOneDrivePath, vbDirectory) <> vbNullString Or InStr(2, strSecPart, pathSep) = 0
                        strSecPart = Mid$(strSecPart, InStr(2, strSecPart, pathSep))
                        GetLocalOneDrivePath = strMountpoint & strSecPart
                    Loop
                    Dim fso As New FileSystemObject
                    If Not fso.FolderExists(GetLocalOneDrivePath) Then GetLocalOneDrivePath = strPath 'OneDrive folder excluded from sync (@6DiegoDiego9)
                    Exit Function
                End If
            Next subKey
        End If
    End If
        
    GetLocalOneDrivePath = strPath 'pass unchanged
End Function

Private Function IsPathRelative(ByVal sPath As String) As Boolean
    'PathIsRelative interprets a properly formed url as relative, so add a check for url too
    If PathIsRelative(sPath) = 1 And PathIsURL(sPath) = 0 Then IsPathRelative = True Else IsPathRelative = False
End Function

Private Function IsPathHTTPS(ByVal sPath As String) As Boolean
    If VBA.Left$(sPath, 8) = "https://" Then IsPathHTTPS = True Else IsPathHTTPS = False
End Function

Private Function IsPathUrl(ByVal sPath As String) As Boolean
    If PathIsURL(sPath) = 1 Then IsPathUrl = True Else IsPathUrl = False
End Function

Private Function IsArrayInitialized(ByRef arry() As Variant) As Boolean
    If (Not arry) = -1 Then IsArrayInitialized = False Else IsArrayInitialized = True
End Function

Public Function GetBrowserName(ByVal browser As svbaBrowser) As String
    Select Case browser
    Case svbaBrowser.Chrome
        GetBrowserName = "chrome"
    Case svbaBrowser.Edge
        GetBrowserName = "msedge"
    Case svbaBrowser.Firefox
        GetBrowserName = "firefox"
    Case svbaBrowser.IE
        GetBrowserName = "internet explorer"
    End Select
End Function

#If ExcelHost Then

Private Function ActiveVBAProjectFolderPath() As String
    'This returns the calling code project's parent document path. So if caller is from a project that references the SeleniumVBA Add-in
    'then this returns the path to the caller, not the Add-in (unless they are the same).
    'But be aware that if qc'ing this routine in Debug mode, the path to this SeleniumVBA project will be returned, which
    'may not be the caller's intended target if it resides in a different project.

    Dim sRespType As String
    sRespType = TypeName(Application.Caller)
    If sRespType <> "Error" Then 'eg. if launched by a formula or a shape button in a worksheet
        ActiveVBAProjectFolderPath = Application.ActiveWorkbook.Path
    Else 'if launched in the VBE
        If VBAIsTrusted Then
            Dim fso As New FileSystemObject
            'below will return an error if active project's host doc has not yet been saved, even if access trusted
            On Error Resume Next
            ActiveVBAProjectFolderPath = fso.GetParentFolderName(Application.VBE.ActiveVBProject.fileName)
            On Error GoTo 0
        Else 'if Excel security setting "Trust access to the VBA project object model" is not enabled
            Dim ThisAppProcessID As Long
            GetWindowThreadProcessId Application.hWnd, ThisAppProcessID
            Do 'search for this VBE window
                Dim hWnd As LongPtr
                hWnd = FindWindowEx(0, hWnd, "wndclass_desked_gsk", vbNullString)
                If hWnd > 0 Then
                    Dim WndProcessID As Long
                    GetWindowThreadProcessId hWnd, WndProcessID
                    If ThisAppProcessID = WndProcessID Then
                        'get its caption
                        Dim Length As Long, caption As String, result As Long
                        Length = GetWindowTextLength(hWnd)
                        caption = Space$(Length + 1)
                        result = GetWindowText(hWnd, caption, Length + 1)
                        
                        'extract filename from the caption
                        Dim oRegex As New RegExp
                        oRegex.Pattern = "^Microsoft Visual Basic[^-]* - ([^[]*) \[[^]]+\] - \[[^(]+\([^)]+\)\]$"
                        Dim regexRes As MatchCollection
                        Set regexRes = oRegex.Execute(Left$(caption, result))
                        If regexRes.Count = 1 Then
                            Dim sFilename As String
                            sFilename = regexRes.Item(0).SubMatches(0)
                            'this returns vbNullString if workbook has not been saved yet
                            ActiveVBAProjectFolderPath = Application.Workbooks(sFilename).Path
                        Else
                            Err.raise 1, , "Error: unable to extract filename from VBE window caption. Check the extraction regex."
                        End If
                    End If
                End If
            Loop Until hWnd = 0
        End If
    End If
    If ActiveVBAProjectFolderPath = vbNullString Then Err.raise 1, , "Error: unable to get the active VBProject path - make sure the parent document has been saved."
End Function

Public Function ThisLibFolderPath() As String
    'returns the path of this library - not the path of the active vba project, which may be referencing this library
    ThisLibFolderPath = Application.ThisWorkbook.Path
End Function

Private Function VBAIsTrusted() As Boolean
    VBAIsTrusted = False
    On Error Resume Next
    VBAIsTrusted = (Application.VBE.VBProjects.Count) > 0
    On Error GoTo 0
End Function

#ElseIf AccessHost Then

Private Function ActiveVBAProjectFolderPath() As String
    'This returns the calling code project's parent document path. So if caller is from a project that references the SeleniumVBA Add-in
    'then this returns the path to the caller, not the Add-in (unless they are the same).
    'But be aware that if qc'ing this routine in Debug mode, the path to this SeleniumVBA project will be returned, which
    'may not be the caller's intended target if it resides in a different project.
    Dim strPath As String
    
    strPath = vbNullString
    
    'if the parent document holding the active vba project has not yet been saved, then Application.VBE.ActiveVBProject.Filename
    'will throw an error so trap and report below...
    
    On Error Resume Next
    strPath = Application.VBE.ActiveVBProject.fileName
    On Error GoTo 0
    
    If strPath <> vbNullString Then
        Dim fso As New FileSystemObject
        strPath = fso.GetParentFolderName(strPath)
        ActiveVBAProjectFolderPath = strPath
    Else
        Err.raise 1, "WebShared", "Error: Attempting to reference a folder/file path relative to the parent document location of this active code project - save the parent document first."
    End If
End Function

Public Function ThisLibFolderPath() As String
    'returns the path of this library - not the path of the active vba project, which may be referencing this library"
    ThisLibFolderPath = Application.CodeProject.Path
End Function

#End If

Private Function ExpandEnvironVariable(ByVal inputPath As String) As String
    'this searches input path for %[Environ Variable]% pattern and if found, then replaces with the path equivalent
    Dim wsh As New IWshRuntimeLibrary.WshShell
    ExpandEnvironVariable = wsh.ExpandEnvironmentStrings(inputPath)
End Function

Public Function ReadIniFileEntry(ByVal filePath As String, ByVal section As String, ByVal keyName As String, Optional ByVal defaultValue As Variant = vbNullString) As String
    'reads a single settings file entry
    Const lenStr = 255
    Dim outputLen As Long
    Dim retStr As String * lenStr
    Dim fso As New FileSystemObject
    
    'check if optional settinsg file exists - if not then use default and exit
    If Not fso.FileExists(filePath) Then
        ReadIniFileEntry = defaultValue
        Exit Function
    End If
    
    'try to read and return the section/keyName value - if not then use default and exit
    retStr = Space(lenStr)
    outputLen = GetPrivateProfileString(section, keyName, vbNullString, retStr, lenStr, filePath)
    If outputLen Then
        ReadIniFileEntry = Left$(retStr, outputLen)
    Else
        ReadIniFileEntry = defaultValue
    End If
End Function

Public Function EnumTextToValue(ByVal enumText As String) As Long
    'this function converts an enum string read from the settings file to it's corresponding enum value
    enumText = Trim$(enumText)
    If IsNumeric(enumText) Then
        EnumTextToValue = VBA.val(enumText)
        Exit Function
    End If
    Select Case LCase(enumText)
    Case LCase("svbaNotCompatible")
        EnumTextToValue = svbaCompatibility.svbaNotCompatible
    Case LCase("svbaMajor")
        EnumTextToValue = svbaCompatibility.svbaMajor
    Case LCase("svbaMinor")
        EnumTextToValue = svbaCompatibility.svbaMinor
    Case LCase("svbaBuildMajor")
        EnumTextToValue = svbaCompatibility.svbaBuildMajor
    Case LCase("svbaExactMatch")
        EnumTextToValue = svbaCompatibility.svbaExactMatch
    Case LCase("vbHide")
        EnumTextToValue = VbAppWinStyle.vbHide
    Case LCase("vbMaximizedFocus")
        EnumTextToValue = VbAppWinStyle.vbMaximizedFocus
    Case LCase("vbMinimizedFocus")
        EnumTextToValue = VbAppWinStyle.vbMinimizedFocus
    Case LCase("vbMinimizedNoFocus")
        EnumTextToValue = VbAppWinStyle.vbMinimizedNoFocus
    Case LCase("vbNormalFocus")
        EnumTextToValue = VbAppWinStyle.vbNormalFocus
    Case LCase("vbNormalNoFocus")
        EnumTextToValue = VbAppWinStyle.vbNormalNoFocus
    Case LCase("svbaLandscape")
        EnumTextToValue = svbaOrientation.svbaLandscape
    Case LCase("svbaPortrait")
        EnumTextToValue = svbaOrientation.svbaPortrait
    Case LCase("svbaCentimeters")
        EnumTextToValue = svbaUnits.svbaCentimeters
    Case LCase("svbaInches")
        EnumTextToValue = svbaUnits.svbaInches
    Case Else
        Err.raise 1, "WebShared", "Settings file enum value " & enumText & " not recognized"
    End Select
End Function

Public Sub Sleep(ByVal ms As Currency)
'Enhanced sleep proc. featuring <0.0% CPU usage, DoEvents, accuracy +-<10ms
'Better Sleep proc. featuring <0.0% CPU usage, DoEvents, accuracy +-<10ms
'Uses "Currency" as a good-enough workaround to avoid the complexity of LARGE_INTEGER (see https://stackoverflow.com/a/31387007)
'Note: VBA.Timer ( + VBA.Date for midnight adjustment) and VBA.Now avoided for accuracy issues (10-15ms and occasionally even worse? see https://stackoverflow.com/questions/68767198/is-this-unstable-vba-timer-behavior-real-or-am-i-doing-something-wrong)
    Dim cTimeStart As Currency, cTimeEnd As Currency
    Dim dTimeElapsed As Currency, cTimeTarget As Currency
    Dim cApproxDelay As Currency
    
    getTime cTimeStart
    
    Static cPerSecond As Currency
    If cPerSecond = 0 Then getFrequency cPerSecond
    cTimeTarget = ms * (cPerSecond / 1000)
    
    If ms <= 25 Then
        'empty loop for improved accuracy (SleepWinAPI alone costs 2-15ms and DoEvents 2-8ms)
        Do
            getTime cTimeEnd
        Loop Until cTimeEnd - cTimeStart >= cTimeTarget
        Exit Sub
    Else 'fully featured loop
        SleepWinAPI 5 '"WaitMessage" avoided because it costs 0.0* to 2**(!) ms
        DoEvents
        getTime cTimeEnd
        cApproxDelay = (cTimeEnd - cTimeStart) / 2
        
        cTimeTarget = cTimeTarget - cApproxDelay
        Do While (cTimeEnd - cTimeStart) < cTimeTarget
            SleepWinAPI 1
            DoEvents
            getTime cTimeEnd
        Loop
    End If
End Sub

