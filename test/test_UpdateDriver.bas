Attribute VB_Name = "test_UpdateDriver"

Sub test_AutoUpdateDrivers()
    '
    'GCUser99: I set the default in this version of SeleniumVBA NOT to check for Driver/Browser version alignment
    'due to lack of testing thus far but it works for me and so....
    '
    'if user wants SeleniumVBA to automatically check the Selenium WebdDriver and Browser versions for compatibility,
    'and if not compatible, then to automatically download and install drivers, then user must modify the following
    'line in WebDriver class:
    '
    'Private Const CheckDriverBrowserVersionAlignment = False
    '
    'To:
    '
    'Private Const CheckDriverBrowserVersionAlignment = True
    '
    Dim Driver As New WebDriver
    
    Driver.Chrome 'here is where the version alignment checking and fixing happens - you will be prompted to install WebDriver if needed
    Driver.Edge 'here is where the version alignment checking and fixing happens - you will be prompted to install WebDriver if needed
    
    Driver.Shutdown

End Sub


Sub test_UpdateSeleniumBasic()
    'this is for Florent Breheret's SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue for writing to this directory so you may need to run as administrator
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



