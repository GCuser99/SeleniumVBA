Attribute VB_Name = "test_UpdateDriver"
'IMPORTANT!!!!
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

Sub test_updateDrivers()
    'this checks if driver is installed, or if installed driver is compatibile
    'with installed browser, and then if needed, installs an updated driver
    Dim mngr As New WebDriverManager
    
    MsgBox mngr.AlignEdgeDriverWithBrowser(), , "SeleniumVBA"
    MsgBox mngr.AlignChromeDriverWithBrowser(), , "SeleniumVBA"
    MsgBox mngr.AlignFirefoxDriverWithBrowser(), , "SeleniumVBA"
End Sub

Sub test_updateDriversForSeleniumBasic()
    'this is for Florent Breheret's SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue for writing to this directory so you may have to run as administrator
    Dim mngr As New WebDriverManager
    
    mngr.DefaultBinaryFolder = mngr.GetSeleniumBasicFolder
    
    MsgBox mngr.AlignEdgeDriverWithBrowser("edgedriver.exe"), , "SeleniumVBA"
    MsgBox mngr.AlignChromeDriverWithBrowser("chromedriver.exe"), , "SeleniumVBA"
    MsgBox mngr.AlignFirefoxDriverWithBrowser("geckodriver.exe"), , "SeleniumVBA"
End Sub
