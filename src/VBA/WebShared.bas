Attribute VB_Name = "WebShared"
'@folder("SeleniumVBA.Source")
' ==========================================================================
' SeleniumVBA v6.3
'
' A Selenium wrapper for browser automation developed for MS Office VBA
'
' https://github.com/GCuser99/SeleniumVBA/tree/main
'
' Contact Info:
'
' https://github.com/6DiegoDiego9
' https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2023-2024, GCUser99 and 6DiegoDiego9 (https://github.com/GCuser99/SeleniumVBA)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================
' For more info:
' https://docs.microsoft.com/en-us/dotnet/standard/io/file-path-formats
' http://vbnet.mvps.org/index.html?code/fileapi/pathisrelative.htm
' https://stackoverflow.com/questions/57475738/ (for use of SetCurrentDirectory)
' https://stackoverflow.com/a/72736800/11738627 (handling of OneDrive/SharePoint cloud urls)

Option Explicit
Option Private Module

'for the Sleep procedure
Public Declare PtrSafe Sub SleepWinAPI Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
Public Declare PtrSafe Function GetFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Public Declare PtrSafe Function GetTime Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef counter As Currency) As Long

Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryW" (ByVal lpPathName As LongPtr) As Long
Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeW" (ByVal pszPath As LongPtr) As Long
Private Declare PtrSafe Function PathIsURL Lib "shlwapi" Alias "PathIsURLW" (ByVal pszPath As LongPtr) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpDefault As LongPtr, lpReturnedString As Any, ByVal nSize As Long, ByVal lpFilename As LongPtr) As Long

Public Declare PtrSafe Function UrlDownloadToFile Lib "urlmon" Alias "URLDownloadToFileW" (ByVal pCaller As Long, ByVal szURL As LongPtr, ByVal szFileName As LongPtr, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptbinarytostringw
Private Declare PtrSafe Function CryptBinaryToStringW Lib "crypt32" (pbBinary As Any, ByVal cbBinary As Long, ByVal dwFlags As CRYPT_STRING_OPTIONS, ByVal pszString As LongPtr, pcchString As Long) As Long
'https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptstringtobinaryw
Private Declare PtrSafe Function CryptStringToBinaryW Lib "crypt32" (ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As CRYPT_STRING_OPTIONS, pbBinary As Any, pcbBinary As Long, Optional pdwSkip As Long, Optional pdwFlags As CRYPT_STRING_OPTIONS) As Long

Private Enum CRYPT_STRING_OPTIONS
    CRYPT_STRING_BASE64 = &H1           'Base64, without headers.
    CRYPT_STRING_NOCRLF = &H40000000    'Do not append any new line characters to the encoded string.
    CRYPT_STRING_NOCR = &H80000000      'this will use vbLF for new line character (as opposed to default vbCrLf), compatible with MSXML2
End Enum

Public Function getFullLocalPath(ByVal inputPath As String, Optional ByVal basePath As String = vbNullString) As String
    'Returns an absolute path from a relative path and a fully qualified base path.
    'basePath defaults to the folder path of the document that holds the Active VBA Project
    'fso.GetAbsolutePathName interprets a url as a relative path, so must avoid for url's

    Dim fso As New Scripting.FileSystemObject, savePath As String

    'make sure no rogue beginning or ending spaces and expand "%[Environ Variable]%" if in the path
    inputPath = expandEnvironVariable(VBA.Trim$(inputPath))

    If Not isPathRelative(inputPath) Then 'its an absolute path
        'just in case OneDrive/SharePoint user has specified a path built with ThisWorkbook.Path...
        If isPathHTTPS(inputPath) Then inputPath = getLocalOneDrivePath(inputPath)
        
        'normalize the path if its not a url - this insures that path separators are correct, and
        'if a folder, has no ending separator
        If Not isPathUrl(inputPath) Then inputPath = fso.GetAbsolutePathName(inputPath)
        
        getFullLocalPath = inputPath
    Else 'ok then convert relative path to absolute
        'make sure no unintended beginning or ending spaces
        basePath = expandEnvironVariable(VBA.Trim$(basePath))
        
        If basePath = vbNullString Then
            basePath = activeVBAProjectFolderPath
        Else
            'it's possible that user specified a relative reference folder path - convert it to absolute relative to
            'the folder path of the document that holds the Active VBA Project
            If isPathRelative(basePath) Then basePath = getFullLocalPath(basePath, activeVBAProjectFolderPath)
        End If

        'convert OneDrive path if needed
        If isPathHTTPS(basePath) Then basePath = getLocalOneDrivePath(basePath)
        
        'check that reference path exists and notify user if not
        If Not fso.FolderExists(basePath) Then
            If Not isPathHTTPS(basePath) Then 'its a url which fso doesn't support - must trust that it exists (@6DiegoDiego9)
                Err.Raise 1, "WebShared", "Reference folder basePath does not exist." & vbNewLine & vbNewLine & basePath & vbNewLine & vbNewLine & "Please specify a valid folder path."
            End If
        End If
        
        'employ fso to make the conversion of relative path to absolute
        savePath = CurDir$()
        SetCurrentDirectory StrPtr(basePath)
        getFullLocalPath = fso.GetAbsolutePathName(inputPath)
        SetCurrentDirectory StrPtr(savePath)
    End If
End Function

Private Function getLocalOneDrivePath(ByVal strPath As String) As String
    'for more info, see https://stackoverflow.com/a/72736800/11738627 post by Guido Witt-Dorring
    'this function returns the original/local disk path associated with a synched OneDrive or SharePoint cloud url
    
    If isPathHTTPS(strPath) Then
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
        
        If isArrayInitialized(subKeys) Then 'found OneDrive in registry
            For Each subKey In subKeys
                objReg.GetStringValue HKEY_CURRENT_USER, regPath & subKey, "UrlNamespace", strValue
                If InStr(strPath, strValue) > 0 Then
                    objReg.GetStringValue HKEY_CURRENT_USER, regPath & subKey, "MountPoint", strMountpoint
                    strSecPart = Replace$(Mid$(strPath, Len(strValue)), "/", pathSep)
                    getLocalOneDrivePath = strMountpoint & strSecPart
        
                    Do Until Dir(getLocalOneDrivePath, vbDirectory) <> vbNullString Or InStr(2, strSecPart, pathSep) = 0
                        strSecPart = Mid$(strSecPart, InStr(2, strSecPart, pathSep))
                        getLocalOneDrivePath = strMountpoint & strSecPart
                    Loop
                    Dim fso As New FileSystemObject
                    If Not fso.FolderExists(getLocalOneDrivePath) Then getLocalOneDrivePath = strPath 'OneDrive folder excluded from sync (@6DiegoDiego9)
                    Exit Function
                End If
            Next subKey
        End If
    End If
        
    getLocalOneDrivePath = strPath 'pass unchanged
End Function

Private Function isPathRelative(ByVal sPath As String) As Boolean
    'PathIsRelative interprets a properly formed url as relative, so add a check for url too
    isPathRelative = (PathIsRelative(StrPtr(sPath)) = 1 And PathIsURL(StrPtr(sPath)) = 0)
End Function

Private Function isPathHTTPS(ByVal sPath As String) As Boolean
    isPathHTTPS = (VBA.Left$(sPath, 8) = "https://")
End Function

Private Function isPathUrl(ByVal sPath As String) As Boolean
    isPathUrl = (PathIsURL(StrPtr(sPath)) = 1)
End Function

Private Function isArrayInitialized(ByRef arry() As Variant) As Boolean
    If (Not arry) = -1 Then isArrayInitialized = False Else isArrayInitialized = True
End Function

Private Function activeVBAProjectFolderPath() As String
    'This returns the calling code project's parent document path. If caller is from a project that references the SeleniumVBA Add-in
    'then this returns the path to the caller, not the Add-in (unless they are the same).
    'But be aware that if qc'ing this routine in Debug mode, the path to this SeleniumVBA project will be returned, which
    'may not be the caller's intended target if it resides in a different project.
    Dim fso As FileSystemObject
    Dim oApp As Object
    
    Set oApp = Application 'late bound needed for cross-app compatibility
    Select Case oApp.Name
    Case "Microsoft Excel"
        'first check if programatic access to vba is trusted - if so, then this works for all tested cases - then we are done
        If vbaIsTrusted Then
            Set fso = New FileSystemObject
            'below will return an error if active project's host doc has not yet been saved (has no valid path yet)
            On Error Resume Next
            activeVBAProjectFolderPath = fso.GetParentFolderName(oApp.VBE.ActiveVBProject.fileName)
            If Err.Number > 0 Then Err.Raise 1, , "Error: unable to get the active Caller path - make sure the parent document has been saved."
            On Error GoTo 0
            Exit Function
        End If
        
        'handle vba-is-not-trusted cases
        
        'test if procedure was launched by a formula, shape button, or form control (not ActiveX) embedded in a worksheet
        If TypeName(oApp.Caller) <> "Error" Then
            'found embedded forumla, shape, or form control
            activeVBAProjectFolderPath = oApp.ActiveWorkbook.Path
            If activeVBAProjectFolderPath = vbNullString Then Err.Raise 1, , "Error: unable to get the active Caller path - make sure the parent document has been saved."
            Exit Function
        End If
        
        'handle cases where:
        '1) procedure launched in the VBE (Run)
        '2) procedure launched by clicking a worksheet-embedded ActiveX control with VBE window having been opened previously
        Dim ThisAppProcessID As Long
        GetWindowThreadProcessId oApp.hWnd, ThisAppProcessID
        
        Do 'search for an open VBE window
            Dim hWnd As LongPtr
            Dim lpszClass As String
            lpszClass = "wndclass_desked_gsk"
            hWnd = FindWindowEx(0, hWnd, StrPtr(lpszClass), 0&)
            If hWnd > 0 Then
                Dim WndProcessID As Long
                GetWindowThreadProcessId hWnd, WndProcessID
                
                If ThisAppProcessID = WndProcessID Then
                    'get window caption
                    Dim bufferLen As Long, caption As String, result As Long
                    bufferLen = GetWindowTextLength(hWnd)
                    caption = String$(bufferLen + 1, vbNullChar)
                    result = GetWindowText(hWnd, StrPtr(caption), bufferLen + 1)
                    caption = Left$(caption, InStr(caption, vbNullChar) - 1)
                    'extract filename from the caption
                    Dim oRegex As New RegExp
                    oRegex.Pattern = "^Microsoft Visual Basic[^-]*- (.*\.xl\w{1,2})(?:|(?:| -) \[.*\])$"
                    Dim regexRes As MatchCollection
                    Set regexRes = oRegex.execute(caption)
                    If regexRes.Count = 1 Then
                        'found vbe window and succesfully parsed caption
                        Dim sFilename As String
                        sFilename = regexRes.Item(0).SubMatches(0)
                        'the following returns vbNullString if workbook has not been saved (has no valid path yet)
                        activeVBAProjectFolderPath = oApp.Workbooks(sFilename).Path
                        If activeVBAProjectFolderPath = vbNullString Then Err.Raise 1, , "Error: unable to get the active VBProject path - make sure the parent document has been saved."
                        Exit Function
                    Else
                        'handle edge case where embedded ActiveX control is caller and VBE window open but "uninitialized"
                        'this edge case happens under the following conditions:

                        '1) caller is a worksheet-embedded ActiveX button (not a form control)
                        '2) caller workbook is opened, user opens VBE, but then closes without running or editing
                        '3) user clicks on button causing a control event to initiate SeleniumVBA procedure

                        'detect open but "uninitialized" VBE window looking for a pattern like this:
                        'Microsoft Visual Basic for Applications - [Module1 (Code)]
                        oRegex.Pattern = "^Microsoft Visual Basic[^-]*- \[.*\]$"
                        If oRegex.test(caption) Then
                            'found embedded ActiveX control with open but uninitialized VBE window
                            'the uninitialized VBE window was found so done with loop - finish processing this case after exiting loop
                            Exit Do
                        Else
                            Err.Raise 1, , "Error: unable to extract filename from VBE window caption. Please report caption to developers." & vbCrLf & vbCrLf & "Caption:" & vbCrLf & vbCrLf & caption
                        End If
                    End If
                End If
            End If
        Loop Until hWnd = 0

        If oApp.ActiveWorkbook.Path <> vbNullString Then
            'by elimination, user clicked embedded ActiveX control with either not opened or an opened but uninitialized VBE window
            activeVBAProjectFolderPath = oApp.ActiveWorkbook.Path
        Else
            Err.Raise 1, , "Error: unable to get the active Caller path - make sure the parent document has been saved."
        End If
        
    Case "Microsoft Access"
        Dim strPath As String
    
        strPath = vbNullString
        
        'if the parent document holding the active vba project has not yet been saved, then Application.VBE.ActiveVBProject.Filename
        'will throw an error so trap and report below...
        
        On Error Resume Next
        strPath = oApp.VBE.ActiveVBProject.fileName
        On Error GoTo 0
        
        If strPath <> vbNullString Then
            Set fso = New FileSystemObject
            strPath = fso.GetParentFolderName(strPath)
            activeVBAProjectFolderPath = strPath
        Else
            Err.Raise 1, "WebShared", "Error: Attempting to reference a folder/file path relative to the parent document location of this active code project - save the parent document first."
        End If
    Case Else
        Err.Raise 1, "WebShared", "Error: Only MS Access and MS Excel supported."
    End Select
End Function

Private Function vbaIsTrusted() As Boolean
    vbaIsTrusted = False
    'Note: this may cause "Run-time Error 1004" if Tools->Options->Error Trapping is set to "Break on All Errors"
    On Error Resume Next
    vbaIsTrusted = (Application.VBE.VBProjects.Count > 0)
    On Error GoTo 0
End Function

Public Function thisLibFolderPath() As String
    'returns the path of this library - not the path of the active vba project, which may be referencing this library"
    Dim oApp As Object
    Set oApp = Application
    Select Case oApp.Name
    Case "Microsoft Excel"
        thisLibFolderPath = oApp.ThisWorkbook.Path
    Case "Microsoft Access"
        thisLibFolderPath = oApp.CodeProject.Path
    Case Else
        Err.Raise 1, "WebShared", "Error: Only MS Access and MS Excel supported."
    End Select
End Function

Public Function getBrowserNameString(ByVal browser As svbaBrowser) As String
    Select Case browser
    Case svbaBrowser.Chrome
        getBrowserNameString = "chrome"
    Case svbaBrowser.Edge
        getBrowserNameString = "msedge"
    Case svbaBrowser.Firefox
        getBrowserNameString = "firefox"
    End Select
End Function

Private Function expandEnvironVariable(ByVal inputPath As String) As String
    'this searches input path for %[Environ Variable]% pattern and if found, then replaces with the path equivalent
    Dim wsh As New IWshRuntimeLibrary.WshShell
    expandEnvironVariable = wsh.ExpandEnvironmentStrings(inputPath)
End Function

Public Function readIniFileEntry(ByVal filePath As String, ByVal section As String, ByVal keyName As String, Optional ByVal defaultValue As Variant = vbNullString, Optional ByVal useDefaultValue As Boolean = False) As String
    'reads a single settings file entry
    Const lenBuffer = 512
    Dim outputLen As Long
    Dim fso As FileSystemObject
    Dim buffer() As Byte
    
    If useDefaultValue Then 'quick escape!
        readIniFileEntry = defaultValue
        Exit Function
    End If
    
    'check if optional settings file exists - if not then use default and exit
    Set fso = New FileSystemObject
    If Not fso.FileExists(filePath) Then
        readIniFileEntry = defaultValue
        Exit Function
    End If
    
    'try to read and return the section/keyName value - if not then use default and exit

    ReDim buffer(0 To lenBuffer - 1)

    outputLen = GetPrivateProfileString( _
                    StrPtr(section), _
                    StrPtr(keyName), _
                    0&, _
                    buffer(0), _
                    lenBuffer, _
                    StrPtr(filePath))
                    
    If outputLen Then
        readIniFileEntry = Left$(buffer, outputLen)
    Else
        readIniFileEntry = defaultValue
    End If
End Function

Public Function enumTextToValue(ByVal enumText As String) As Long
    'this function converts an enum string read from the settings file to it's corresponding enum value
    enumText = Trim$(enumText)
    If IsNumeric(enumText) Then
        enumTextToValue = VBA.val(enumText)
        Exit Function
    End If
    Select Case LCase$(enumText)
    Case "svbanotcompatible": enumTextToValue = svbaCompatibility.svbaNotCompatible
    Case "svbamajor": enumTextToValue = svbaCompatibility.svbaMajor
    Case "svbaminor": enumTextToValue = svbaCompatibility.svbaMinor
    Case "svbabuildmajor": enumTextToValue = svbaCompatibility.svbaBuildMajor
    Case "svbaexactmatch": enumTextToValue = svbaCompatibility.svbaExactMatch
    Case "vbhide": enumTextToValue = VbAppWinStyle.vbHide
    Case "vbmaximizedfocus": enumTextToValue = VbAppWinStyle.vbMaximizedFocus
    Case "vbminimizedfocus": enumTextToValue = VbAppWinStyle.vbMinimizedFocus
    Case "vbminimizednofocus": enumTextToValue = VbAppWinStyle.vbMinimizedNoFocus
    Case "vbnormalfocus": enumTextToValue = VbAppWinStyle.vbNormalFocus
    Case "vbnormalnofocus": enumTextToValue = VbAppWinStyle.vbNormalNoFocus
    Case "svbalandscape": enumTextToValue = svbaOrientation.svbaLandscape
    Case "svbaportrait": enumTextToValue = svbaOrientation.svbaPortrait
    Case "svbacentimeters": enumTextToValue = svbaUnits.svbaCentimeters
    Case "svbainches": enumTextToValue = svbaUnits.svbaInches
    Case Else: Err.Raise 1, "WebShared", "Settings file enum value " & enumText & " not recognized"
    End Select
End Function

Public Sub sleep(ByVal ms As Currency)
    'Enhanced sleep proc. featuring <0.0% CPU usage, DoEvents, precision +-<10ms
    'Better Sleep proc. featuring <0.0% CPU usage, DoEvents, precision +-<10ms
    'Uses "Currency" as a good-enough workaround to avoid the complexity of a LARGE_INTEGER (see https://stackoverflow.com/a/31387007)
    'Note: VBA.Timer ( + VBA.Date for midnight adjustment) and VBA.Now avoided for accuracy issues (10-15ms and occasionally even worse? see https://stackoverflow.com/questions/68767198/is-this-unstable-vba-timer-behavior-real-or-am-i-doing-something-wrong)
    Dim cTimeStart As Currency, cTimeEnd As Currency
    Dim dTimeElapsed As Currency, cTimeTarget As Currency
    Dim cApproxDelay As Currency
    
    GetTime cTimeStart
    
    Static cPerSecond As Currency
    If cPerSecond = 0 Then GetFrequency cPerSecond
    cTimeTarget = ms * (cPerSecond / 1000)
    
    If ms <= 25 Then
        'empty loop for improved accuracy (SleepWinAPI alone costs 2-15ms and DoEvents 2-8ms)
        Do
            GetTime cTimeEnd
        Loop Until cTimeEnd - cTimeStart >= cTimeTarget
        Exit Sub
    Else 'fully featured loop
        SleepWinAPI 5 '"WaitMessage" avoided because it costs 0.0* to 2**(!) ms
        DoEvents
        GetTime cTimeEnd
        cApproxDelay = (cTimeEnd - cTimeStart) / 2
        
        cTimeTarget = cTimeTarget - cApproxDelay
        Do While (cTimeEnd - cTimeStart) < cTimeTarget
            SleepWinAPI 1
            DoEvents
            GetTime cTimeEnd
        Loop
    End If
End Sub

Public Function Max(ParamArray numberList() As Variant) As Variant
    Dim i As Long
    Max = numberList(LBound(numberList))
    For i = LBound(numberList) + 1 To UBound(numberList)
        If numberList(i) > Max Then Max = numberList(i)
    Next i
End Function

Public Function AppActivate(ByVal partialWindowText As String, Optional ByVal Wait As Boolean = False) As Boolean
    'The VBA.AppActivate throws an error if a match is not found
    'This function wraps VBA's AppActivate but returns a boolean signifying whether the match was found
    'which allows for looping until found
    On Error GoTo winNotFound
    VBA.AppActivate partialWindowText, Wait
    AppActivate = True
    Exit Function
winNotFound:
    AppActivate = False
End Function

Public Function IsActiveWindowVBIDE() As Boolean
    'determines whether the VBDIDE is the active window or not
    Dim winTitle As String
    winTitle = String$(200, vbNullChar)
    GetWindowText GetForegroundWindow(), StrPtr(winTitle), 200
    IsActiveWindowVBIDE = Left$(winTitle, InStr(winTitle, vbNullChar) - 1) Like "Microsoft Visual Basic for Applications -*"
End Function

Public Function splitKeyString(ByVal keys As String) As Collection
    'splits the input string into a collection of individual characters
    Dim i As Long
    Dim chars As New Collection
    For i = 1 To Len(keys)
        chars.Add Mid$(keys, i, 1)
    Next i
    Set splitKeyString = chars
End Function

Public Function getResponseErrorMessage(resp As Dictionary) As String
    If TypeName(resp("value")) = "Dictionary" Then
        If resp("value").Exists("error") Then getResponseErrorMessage = resp("value")("message")
    End If
End Function

Public Function isResponseError(resp As Dictionary) As Boolean
    isResponseError = False
    If TypeName(resp("value")) = "Dictionary" Then isResponseError = resp("value").Exists("error")
End Function

'this is used to convert an escaped unicode string (e.g. "\u00A9") into the
'single (wide) character equiv. This conversion is required by WebJsonConverter.
Public Function unEscapeUnicode(ByRef keyString As Variant) As String
    Dim oRegExp As New VBScript_RegExp_55.RegExp
    Dim matches As VBScript_RegExp_55.MatchCollection
    Dim match As VBScript_RegExp_55.match
    Dim unEscapedValue As String
    
    oRegExp.Global = True

    'process escaped unicode
    oRegExp.Pattern = "\\u([0-9a-fA-F]{4})"
    Set matches = oRegExp.execute(keyString)
    If matches.Count > 0 Then
        For Each match In matches
            unEscapedValue = ChrW$(val("&H" & match.SubMatches(0)))
            'replace match with unescaped value
            keyString = Replace$(keyString, match.Value, unEscapedValue, Count:=1)
        Next match
    End If
    unEscapeUnicode = keyString
End Function

'this is to replace AscW which outputs negative ascii code for unicode due to its integer return value
Public Function AscWL(ByVal s As String) As Long
    AscWL = CLng(AscW(s)) And &HFFFF&
End Function

Public Function encodeBase64(bytes() As Byte, Optional ByVal useNewLines As Boolean = True, Optional ByVal useCrLfForNewLine As Boolean = True) As String
    'https://gist.github.com/wqweto/0002b7e6c4f92e69c8e8339ed2235b4c
    Dim lSize As Long
    Dim flags As Long
    Dim baseIndex As Long
    Dim numBytes As Long
    
    flags = CRYPT_STRING_BASE64
    If Not useNewLines Then
        flags = flags Or CRYPT_STRING_NOCRLF
    ElseIf Not useCrLfForNewLine Then
        flags = flags Or CRYPT_STRING_NOCR 'compatibility with MSXML2 which uses vbLf
    End If
    
    baseIndex = LBound(bytes)
    numBytes = UBound(bytes) - baseIndex + 1
    
    encodeBase64 = String$(2 * numBytes + 4, 0)
    lSize = Len(encodeBase64) + 1
    Call CryptBinaryToStringW(bytes(baseIndex), numBytes, flags, StrPtr(encodeBase64), lSize)
    encodeBase64 = Left$(encodeBase64, lSize)
End Function

Public Function decodeBase64(sText As String, Optional ByVal baseIndex As Long = 1) As Byte()
    'https://gist.github.com/wqweto/0002b7e6c4f92e69c8e8339ed2235b4c
    Dim lSize As Long
    Dim baOutput() As Byte
    lSize = Len(sText) + 1
    ReDim baOutput(baseIndex To lSize - 1 + baseIndex) As Byte
    If CryptStringToBinaryW(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, baOutput(baseIndex), lSize) <> 0 Then
        ReDim Preserve baOutput(baseIndex To lSize - 1 + baseIndex) As Byte
        decodeBase64 = baOutput
    End If
End Function

Public Sub saveByteArrayToFile(bytearray() As Byte, ByVal filePath As String)
    'we are using ADODB.Stream instead of native Open because later does not support unicode filePaths
    Dim binaryStream As New ADODB.stream
    'specify stream type - we want to save binary data.
    binaryStream.Type = adTypeBinary
  
    'open the stream And write binary data
    binaryStream.Open
    binaryStream.Write bytearray
  
    'save binary data To disk
    binaryStream.SaveToFile filePath, adSaveCreateOverWrite
    binaryStream.Close
End Sub

Public Function readByteArrayFromFile(ByVal filePath As String) As Byte()
    'we are using ADODB.Stream instead of native Open because later does not support unicode filePaths
    Dim binaryStream As New ADODB.stream
    Dim bytearray() As Byte

    'specify stream type - we want to read binary data.
    binaryStream.Type = adTypeBinary
  
    'open the stream
    binaryStream.Open
    
    'load and read binary data from disk
    binaryStream.LoadFromFile filePath
    bytearray = binaryStream.Read 'returns 0-based byte array
  
    binaryStream.Close
    readByteArrayFromFile = bytearray
End Function
