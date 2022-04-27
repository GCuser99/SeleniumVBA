Attribute VB_Name = "test_UpdateDriver"
'ATTENTION!!!!
'---------------------------------------------------------------------------------------------------------------
'
'To immediately check on the version alignment between installed Selenium WebDrivers and Browsers, and to then
'install compatible drivers if not compatible, run the "test_UpdateDriversForSeleniumVBA" subroutine below. This will install the
'compatible versions of WebDriver for both Chrome and Edge, even if you have not yet installed them. Note that
'the default folder for installation is the same folder that this Excel file resides.
'
'---------------------------------------------------------------------------------------------------------------
'
'There is also capability in the WebDriver class to auto-check and conditionally install every time the StartChrome and StartEdge
'methods are invoked. However the default in this version of SeleniumVBA is set NOT to auto-check and install due
'to lack of testing thus far but it works fine for me and so....
'
'SeleniumVBA can auto-check the Selenium WebDriver and Browser versions for compatibility
'when StartChrome and StartEdge methods are invoked, and if not compatible, can automatically download
'and install drivers. To make that happen, modify the following line in WebDriver class:
'
'Private Const CheckDriverBrowserVersionAlignment = False
'
'To:
'
'Private Const CheckDriverBrowserVersionAlignment = True
'
'---------------------------------------------------------------------------------------------------------------
'
'Otherwise if to install the WebDrivers manually, then download from:
'
'Edge: https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
'
'Chrome: https://chromedriver.chromium.org/downloads
'
'Be sure to install drivers with the same major version (number to left of first period)
'as corresponding browser. Place the driver(s) in the same directory as where this Excel file resides.
'
'---------------------------------------------------------------------------------------------------------------

Sub test_UpdateDriversForSeleniumVBA()
    'this checks for installed driver compatibility and then if not, installs updated driver
    Dim mngr As New WebDriverManager
    
    browserName = "msedge"
    driverPath = ".\msedgedriver.exe"
    
    mngr.AlignDriverAndBrowser browserName, driverPath
    
    browserVer = mngr.GetInstalledBrowserVersion(browserName)
    driverVer = mngr.GetInstalledDriverVersion(browserName, driverPath)
    
    MsgBox "Edge " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & browserVer & vbCrLf & "Driver version:    " & driverVer, , "SeleniumVBA"

    browserName = "chrome"
    driverPath = ".\chromedriver.exe"
    
    mngr.AlignDriverAndBrowser browserName, driverPath
    
    browserVer = mngr.GetInstalledBrowserVersion(browserName)
    driverVer = mngr.GetInstalledDriverVersion(browserName, driverPath)
    
    MsgBox "Chrome " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & browserVer & vbCrLf & "Driver version:    " & driverVer, , "SeleniumVBA"
End Sub

Sub test_UpdateDriversForSeleniumBasic()
    'this is for Florent Breheret's SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue for writing to this directory so you may have to run as administrator
    Dim mngr As New WebDriverManager
    
    browserName = "msedge"
    driverPath = mngr.GetSeleniumBasicFolder & "edgedriver.exe"
    
    mngr.AlignDriverAndBrowser browserName, driverPath
    
    browserVer = mngr.GetInstalledBrowserVersion(browserName)
    driverVer = mngr.GetInstalledDriverVersion(browserName, driverPath)
    
    MsgBox "Edge " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & browserVer & vbCrLf & "Driver version:    " & driverVer, , "SeleniumVBA"

    browserName = "chrome"
    driverPath = mngr.GetSeleniumBasicFolder & "chromedriver.exe"
    
    mngr.AlignDriverAndBrowser browserName, driverPath
    
    browserVer = mngr.GetInstalledBrowserVersion(browserName)
    driverVer = mngr.GetInstalledDriverVersion(browserName, driverPath)
    
    MsgBox "Chrome " & "Webdriver and Browser are compatible!" & vbCrLf & vbCrLf & "Browser version: " & browserVer & vbCrLf & "Driver version:    " & driverVer, , "SeleniumVBA"
End Sub
