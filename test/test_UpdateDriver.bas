Attribute VB_Name = "test_UpdateDriver"
'ATTENTION!!!!
'
'The default in this version of SeleniumVBA is set NOT to automaticaaly check for Driver/Browser version alignment
'due to lack of testing thus far but it works for fine me and so....
'
'---------------------------------------------------------------------------------------------------------------
'
'If user wants SeleniumVBA to automatically check the Selenium WebdDriver and Browser versions for compatibility
'when StartChrome and StartEdge methods are invoked, and if not compatible, then to automatically download
'and install drivers, then user must modify the following line in WebDriver class:
'
'Private Const CheckDriverBrowserVersionAlignment = False
'
'To:
'
'Private Const CheckDriverBrowserVersionAlignment = True
'
'---------------------------------------------------------------------------------------------------------------
'
'Otherwise user can install these manually from:
'
'Edge: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
'
'Chrome: https://chromedriver.chromium.org/downloads
'
'Be sure to install drivers with the same major version as corresponding browser.
'Place the driver in the same directory as where this Excel file resides.
'
'---------------------------------------------------------------------------------------------------------------
'
'Another option is to run the subroutine below, which will check compatibility now and install if needed.
'
'---------------------------------------------------------------------------------------------------------------

Sub test_UpdateDriversForSeleniumVBA()
    Dim wdmgr As New WebDriverManager
    
    driverPath = ".\msedgedriver.exe"
    browserName = "msedge"
    
    If Not wdmgr.IsInstalledDriverCompatible(browserName, , driverPath) Then
        resp = MsgBox("WebDriver is not compatible with installed browser - would you like to install the compatible WebDriver?", vbYesNo)
        If resp = vbYes Then
            bverInstalled = wdmgr.GetInstalledBrowserVersion(browserName)
            dverCompat = wdmgr.GetCompatibleDriverVersion(browserName, bverInstalled)
            wdmgr.DownloadAndInstall browserName, dverCompat, driverPath
            MsgBox "Edge" & " " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & bverInstalled & vbCrLf & "Driver version:    " & dverCompat, , "SeleniumVBA"
        End If
    Else
        MsgBox "Edge " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & wdmgr.GetInstalledBrowserVersion(browserName) & vbCrLf & "Driver version:    " & wdmgr.GetInstalledDriverVersion(browserName), , "SeleniumVBA"
    End If

    driverPath = ".\chromedriver.exe"
    browserName = "chrome"
    
    If Not wdmgr.IsInstalledDriverCompatible(browserName, , driverPath) Then
        resp = MsgBox("WebDriver is not compatible with installed browser - would you like to install the compatible WebDriver?", vbYesNo)
        If resp = vbYes Then
            bverInstalled = wdmgr.GetInstalledBrowserVersion(browserName)
            dverCompat = wdmgr.GetCompatibleDriverVersion(browserName, bverInstalled)
            wdmgr.DownloadAndInstall browserName, dverCompat, driverPath
            MsgBox "Chrome" & " " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & bverInstalled & vbCrLf & "Driver version:    " & dverCompat, , "SeleniumVBA"
        End If
    Else
        MsgBox "Chrome " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & wdmgr.GetInstalledBrowserVersion(browserName) & vbCrLf & "Driver version:    " & wdmgr.GetInstalledDriverVersion(browserName), , "SeleniumVBA"
    End If

End Sub

Sub test_UpdateDriversForSeleniumBasic()
    'this is for Florent Breheret's SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue for writing to this directory so you may need to run as administrator
    Dim wdmgr As New WebDriverManager
    
    bname = "msedge"
    
    If Not wdmgr.IsInstalledDriverCompatible(bname, , wdmgr.GetSeleniumBasicFolder & "\" & "edgedriver.exe") Then
        Debug.Print "installing latest driver"
        bverInstalled = wdmgr.GetInstalledBrowserVersion(bname)
        dverCompat = wdmgr.GetCompatibleDriverVersion(bname, bverInstalled)
        wdmgr.DownloadAndInstall bname, dverCompat, wdmgr.GetSeleniumBasicFolder & "\" & "edgedriver.exe"
    End If
    
    bverInstalled = wdmgr.GetInstalledBrowserVersion(bname)
    
    If bverInstalled = "browser not installed" Then MsgBox "browser not installed": Exit Sub
    
    dfolder = wdmgr.GetSeleniumBasicFolder
    
    dverinstalled = wdmgr.GetInstalledDriverVersion(bname, dfolder & "\" & "edgedriver.exe")

    dverCompat = wdmgr.GetCompatibleDriverVersion(bname, bverInstalled)
    
    Debug.Print dverCompat, dverinstalled
    
    If dverinstalled <> "driver not found" Then
        clevel = wdmgr.CheckCompatibilityLevel(dverinstalled, dverCompat)

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
        wdmgr.DownloadAndInstall bname, dverCompat, dfolder
    End If
    
    bname = "chrome"
    
    If Not wdmgr.IsInstalledDriverCompatible(bname, , wdmgr.GetSeleniumBasicFolder & "\" & "chromedriver.exe") Then
        Debug.Print "installing latest driver"
        bverInstalled = wdmgr.GetInstalledBrowserVersion(bname)
        dverCompat = wdmgr.GetCompatibleDriverVersion(bname, bverInstalled)
        wdmgr.DownloadAndInstall bname, dverCompat, wdmgr.GetSeleniumBasicFolder & "\" & "chromedriver.exe"
    End If
    
    bverInstalled = wdmgr.GetInstalledBrowserVersion(bname)
    
    If bverInstalled = "browser not installed" Then MsgBox "browser not installed": Exit Sub
    
    dfolder = wdmgr.GetSeleniumBasicFolder
    
    dverinstalled = wdmgr.GetInstalledDriverVersion(bname, dfolder & "\" & "chromedriver.exe")

    dverCompat = wdmgr.GetCompatibleDriverVersion(bname, bverInstalled)
    
    Debug.Print dverCompat, dverinstalled
    
    If dverinstalled <> "driver not found" Then
        clevel = wdmgr.CheckCompatibilityLevel(dverinstalled, dverCompat)

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
        wdmgr.DownloadAndInstall bname, dverCompat, dfolder
    End If

End Sub



