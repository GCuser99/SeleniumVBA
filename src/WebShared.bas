Attribute VB_Name = "WebShared"
' For more info:
' https://docs.microsoft.com/en-us/dotnet/standard/io/file-path-formats
' http://vbnet.mvps.org/index.html?code/fileapi/pathisrelative.htm
' https://stackoverflow.com/questions/57475738/ (for use of SetCurrentDirectory)
' https://stackoverflow.com/a/72736800/11738627 (handling of OneDrive/SharePoint cloud urls)

'Several points of clarification:
'If the basePath is not specified or vbNullString, then the basePath is set to the path of active code project's
'parent document. The user has the ability to change the default basePath through DefaultIOFolder
'The only times in the code base where the default basePath is not specified is in DefaultIOFolder & DefaultDriverFolder

Option Explicit
Option Private Module

Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function PathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal pszPath As String) As Long

Public Function GetFullLocalPath(ByVal inputPath As String, Optional ByVal basePath As String = vbNullString) As String
    'Returns an absolute path from a relative path and a fully qualified base path.
    'basePath defaults to the folder path of the document that holds the Active VBA Project
    'fso.GetAbsolutePathName interprets a url as a relative path, so must avoid for url's

    Dim fso As New Scripting.FileSystemObject, savePath As String

    'make sure no rogue beginning or ending spaces
    inputPath = VBA.Trim(inputPath)

    If Not IsPathRelative(inputPath) Then 'its an absolute path
        'just in case OneDrive/SharePoint user has specified a path built with ThisWorkbook.Path...
        If IsPathHTTPS(inputPath) Then inputPath = GetLocalOneDrivePath(inputPath)
        
        'normalize the path if its not a url - this insures that path separators are correct, and
        'if a folder, has no ending separator
        If Not IsPathUrl(inputPath) Then inputPath = fso.GetAbsolutePathName(inputPath)
        
        GetFullLocalPath = inputPath
    Else 'ok then convert relative path to absolute
        'make sure no unintended beginning or ending spaces
        basePath = VBA.Trim(basePath)
        
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
            Err.raise 1, "WebShared", "Reference folder basePath does not exist." & vbNewLine & vbNewLine & basePath & vbNewLine & vbNewLine & "Please specify a valid folder path."
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
                    strSecPart = Replace(Mid(strPath, Len(strValue)), "/", pathSep)
                    GetLocalOneDrivePath = strMountpoint & strSecPart
        
                    Do Until Dir(GetLocalOneDrivePath, vbDirectory) <> vbNullString Or InStr(2, strSecPart, pathSep) = 0
                        strSecPart = Mid(strSecPart, InStr(2, strSecPart, pathSep))
                        GetLocalOneDrivePath = strMountpoint & strSecPart
                    Loop
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

Private Function ActiveVBAProjectFolderPath() As String
    'This returns the calling code project's parent document path. So if caller is from a project that references the SeleniumVBA Add-in
    'then this returns the path to the caller, not the Add-in (unless they are the same).
    'But be aware that if qc'ing this routine in Debug mode, the path to this SeleniumVBA project will be returned, which
    'may not be the caller's intended target if it resides in a different project.
    Dim fso As New FileSystemObject
    Dim strPath As String
    
    strPath = vbNullString
    'if the parent document holding the active vba project has not yet been saved, then Application.VBE.ActiveVBProject.Filename
    'will throw an error so trap and report below...
    On Error Resume Next
    strPath = Application.VBE.ActiveVBProject.Filename
    On Error GoTo 0
    If strPath <> vbNullString Then
        strPath = fso.GetParentFolderName(strPath)
        ActiveVBAProjectFolderPath = strPath
    Else
        Err.raise 1, "WebShared", "Error: Attempting to reference a folder/file path relative to the parent document location of this active code project - save the parent document first."
    End If
End Function
