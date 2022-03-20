Attribute VB_Name = "test_Update_Driver"

Sub test_UpdateSeleniumBasic()
    'this is for SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue - need to work on it
    Dim wdmgr As New WebDriverManager
    
    bname = "chrome" 'or "msedge"
    
    If Not wdmgr.IsInstalledDriverCompatible(bname, , wdmgr.GetSeleniumBasicFolder & "\" & "edgedriver.exe") Then
        Debug.Print "installing latest driver"
        bverinstalled = wdmgr.GetInstalledBrowserVersion(bname)
        dvercompat = wdmgr.GetCompatibleDriverVersion(bname, bverinstalled)
        wdmgr.DownloadAndInstall bname, dvercompat, wdmgr.GetSeleniumBasicFolder & "\" & "edgedriver.exe"
    End If
    
    bverinstalled = wdmgr.GetInstalledBrowserVersion(bname)
    
    If bverinstalled = "browser not installed" Then MsgBox "browser not installed": Exit Sub
    
    dfolder = wdmgr.GetSeleniumBasicFolder
    
    dverinstalled = wdmgr.GetInstalledDriverVersion(bname, dfolder)

    dvercompat = wdmgr.GetCompatibleDriverVersion(bname, bverinstalled)
    
    Debug.Print dvercompat, dverinstalled
    
    'file not found could be a compatibility of -1
    If dverinstalled <> "driver not found" Then
        clevel = wdmgr.CheckCompatibilityLevel(dverinstalled, dvercompat)

        Select Case clevel
        Case 0
            updateresp = MsgBox("The browser and WebDriver are incompatible - would you like to update the WebDriver now?", vbYesNo)
        Case 1, 2
            updateresp = MsgBox("The browser and WebDriver are compatible but there is a new WebDriver build - would you like to update the WebDriver now?", vbYesNo)
        Case Else
            'minor build version (last two digits) - no need to update
        End Select
    Else
        'show user path here in msgbox
        updateresp = MsgBox("The specified path to WebDriver was not found - would you like to install the WebDriver now?", vbYesNo)
    End If
    
    If updateresp = vbYes Then
        'download and install
        wdmgr.DownloadAndInstall bname, dvercompat, dfolder
    End If
    
End Sub



