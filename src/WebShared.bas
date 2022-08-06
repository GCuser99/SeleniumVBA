Attribute VB_Name = "WebShared"
Option Explicit

Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Public Function GetAbsolutePath(ByVal strPath As String) As String
    Dim fso As New Scripting.FileSystemObject, savePath As String, path As String
    savePath = CurDir()
    path = ThisWorkbook_PathOnDisk
    SetCurrentDirectory path 'VBA ChDrive/ChDir don't work with UNC paths, see https://stackoverflow.com/questions/57475738/
    GetAbsolutePath = fso.GetAbsolutePathName(VBA.Trim(strPath))
    SetCurrentDirectory savePath
End Function

Private Function ThisWorkbook_PathOnDisk() As String
    ' The reason for this function is that when the workbook is opened on a disk synched with OneDrive or SharePoint,
    '  (ThisWorkbook.FullName and) ThisWorkbook.Path returns the correspondent cloud URLs instead than the original path on disk. For example:
    ' "https://d.docs.live.net/e06a[etc...]/MyDocumentFolder/MyFolder"
    ' or "https://mycompany.sharepoint.com/personal/MyName_Company_com/MyDocumentFolder/mycompany/Apps/BlaBla"
    ' causing problems if that path is used with other functions, like ChDrive.
    '
    ' This function must be used as a replacement to "ThisWorkbook.Path" that always returns the original/real path on disk. For example:
    ' "C:\Users\myUserName\OneDrive\Documenti\MyFolder"

    Dim strPath As String
    strPath = ThisWorkbook.path

    If VBA.Left$(strPath, 8) = "https://" Then
        'Original script taken from https://stackoverflow.com/a/72736800/11738627  (credits to GWD and his sources)
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
    
        For Each subKey In subKeys
            objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "UrlNamespace", strValue
            If InStr(strPath, strValue) > 0 Then
                objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, "MountPoint", strMountpoint
                strSecPart = Replace(Mid(strPath, Len(strValue)), "/", pathSep)
                ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
    
                Do Until Dir(ThisWorkbook_PathOnDisk, vbDirectory) <> "" Or InStr(2, strSecPart, pathSep) = 0
                    strSecPart = Mid(strSecPart, InStr(2, strSecPart, pathSep))
                    ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
                Loop
                Exit Function
            End If
        Next subKey
    End If
        
    ThisWorkbook_PathOnDisk = strPath
    
End Function
