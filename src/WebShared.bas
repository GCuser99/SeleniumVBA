Attribute VB_Name = "WebShared"
' For more info:
' https://docs.microsoft.com/en-us/dotnet/standard/io/file-path-formats
' http://vbnet.mvps.org/index.html?code/fileapi/pathisrelative.htm
' https://stackoverflow.com/questions/57475738/ (for use of SetCurrentDirectory)
' https://stackoverflow.com/a/72736800/11738627 (handling of OneDrive/SharePoint cloud urls)

Option Explicit

Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare PtrSafe Function PathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare PtrSafe Function PathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal pszPath As String) As Long

Public Function GetFullLocalPath(ByVal inputPath As String, Optional ByVal basePath As String = "") As String
    'Returns an absolute path from a relative path and a fully qualified base path.
    'basePath defaults to ThisWorkbook.Path
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
        
        If basePath = "" Then basePath = ThisWorkbook.Path
        
        'its possible that user specified a relative reference folder path - convert it to absolute relative to ThisWorkbook.Path
        If IsPathRelative(basePath) Then basePath = GetFullLocalPath(basePath, ThisWorkbook.Path)

        'convert OneDrive path if needed
        If IsPathHTTPS(basePath) Then basePath = GetLocalOneDrivePath(basePath)
        
        'check that reference path exists and notify user if not
        If Not fso.FolderExists(basePath) Then
            Err.raise 1, , "Reference folder basePath does not exist." & vbNewLine & vbNewLine & basePath & vbNewLine & vbNewLine & "Please specify a valid folder path."
        End If
        
        'employ fso to make the conversion of relative path to absolute
        savePath = CurDir()
        SetCurrentDirectory basePath
        GetFullLocalPath = fso.GetAbsolutePathName(inputPath)
        SetCurrentDirectory savePath
    End If
End Function

Private Function GetLocalOneDrivePath(ByVal strPath As String) As String
    ' thanks to @6DiegoDiego9 for doing research on this (see https://stackoverflow.com/a/72736800/11738627)
    ' this function returns the original/local disk path associated with a synched OneDrive or SharePoint cloud url
    
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
        If pathSep = "" Then pathSep = Application.PathSeparator
    
        Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

        regPath = "Software\SyncEngines\Providers\OneDrive\"
        objReg.EnumKey HKEY_CURRENT_USER, regPath, subKeys
        
        If IsArrayInitialized(subKeys) Then 'found OneDrive in registry
            For Each subKey In subKeys
                objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "UrlNamespace", strValue
                If InStr(strPath, strValue) > 0 Then
                    objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "MountPoint", strMountpoint
                    strSecPart = Replace(Mid(strPath, Len(strValue)), "/", pathSep)
                    GetLocalOneDrivePath = strMountpoint & strSecPart
        
                    Do Until Dir(GetLocalOneDrivePath, vbDirectory) <> "" Or InStr(2, strSecPart, pathSep) = 0
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
