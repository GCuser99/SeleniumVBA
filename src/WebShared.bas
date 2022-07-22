Attribute VB_Name = "WebShared"
Option Explicit

Public Function GetAbsolutePath(ByVal strPath As String, Optional ByVal refPath As String = "") As String
    Dim fso As New IWshRuntimeLibrary.FileSystemObject, savepath As String
    savepath = CurDir()
    If refPath = "" Then refPath = ThisWorkbook_PathOnDisk
    ChDrive refPath
    ChDir refPath
    GetAbsolutePath = fso.GetAbsolutePathName(VBA.Trim(strPath))
    ChDrive savepath
    ChDir savepath
End Function

Public Function ThisWorkbook_PathOnDisk() As String
    ' The reason for this function is that when the workbook is opened on a disk synched with OneDrive or SharePoint,
    '  (ThisWorkbook.FullName and) ThisWorkbook.Path returns the correspondent cloud URLs instead than the original path on disk. For example:
    ' "https://d.docs.live.net/e06a[etc...]/MyDocumentFolder/MyFolder"
    ' or "https://mycompany.sharepoint.com/personal/MyName_Company_com/MyDocumentFolder/mycompany/Apps/BlaBla"
    ' causing problems if that path is used with other functions, like ChDrive.
    '
    ' This function must be used as a replacement to "ThisWorkbook.Path" that always returns the original/real path on disk. For example:
    ' "C:\Users\myUserName\OneDrive\Documenti\MyFolder"

    Dim strPath As String
    strPath = ThisWorkbook.Path

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
            objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, _
                                "UrlNamespace", strValue
            If InStr(strPath, strValue) > 0 Then
                objReg.getStringValue HKEY_CURRENT_USER, regPath & subKey, _
                                    "MountPoint", strMountpoint
                strSecPart = Replace(Mid(strPath, Len(strValue)), "/", pathSep)
                ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
    
                Do Until Dir(ThisWorkbook_PathOnDisk, vbDirectory) <> "" Or _
                         InStr(2, strSecPart, pathSep) = 0
                    strSecPart = Mid(strSecPart, InStr(2, strSecPart, pathSep))
                    ThisWorkbook_PathOnDisk = strMountpoint & strSecPart
                Loop
                Exit Function
            End If
        Next
        ThisWorkbook_PathOnDisk = strPath
    Else
        ThisWorkbook_PathOnDisk = strPath
    End If
End Function

Public Function TaskKillbyImage(ByVal taskName As String)
    Dim wsh As New IWshRuntimeLibrary.wshShell
    TaskKillbyImage = wsh.Run("taskkill /f /t /im " & taskName, 0, True)
End Function

Public Function TaskKillbyPid(ByVal pid As String)
    Dim wsh As New IWshRuntimeLibrary.wshShell
   TaskKillbyPid = wsh.Run("taskkill /f /t /pid " & pid, 0, True)
End Function
