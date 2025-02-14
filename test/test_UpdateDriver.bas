Attribute VB_Name = "test_UpdateDriver"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'---------------------------------------------------------------------------------------------------------------
'
'To immediately check on the version alignment between installed Selenium WebDrivers and Browsers, and to then
'install compatible drivers if not compatible, run the "test_UpdateDrivers" subroutine below. This will install the
'compatible versions of WebDriver Chrome, Edge, and Firefox, even if you have not yet installed them. Note that
'the default folder for installation is your Downloads folder.
'
'---------------------------------------------------------------------------------------------------------------
'
'There is also the optional capability in the WebDriver class to auto-check and conditionally install every time the StartChrome, StartEdge,
'and StartFirefox methods are invoked. The default in this version of SeleniumVBA is set to auto-check and install.
'
'If user wishes to turn this functionality off and manage the alignment themselves, then see the Wiki topic
'Advanced Customization at https://github.com/GCuser99/SeleniumVBA/wiki#advanced-customization
'
'---------------------------------------------------------------------------------------------------------------

Sub test_updateDrivers()
    'this checks if driver is installed, or if installed driver is compatibile
    'with installed browser, and then if needed, installs an updated driver
    Dim mngr As SeleniumVBA.WebDriverManager
    
    Set mngr = SeleniumVBA.New_WebDriverManager
    
    'mngr.DefaultDriverFolder = [your binary folder path here] 'defaults to Downloads dir
    
    MsgBox mngr.AlignEdgeDriverWithBrowser()
    MsgBox mngr.AlignChromeDriverWithBrowser()
    MsgBox mngr.AlignFirefoxDriverWithBrowser()
End Sub

Sub test_updateDriversForSeleniumBasic()
    'this is for Florent Breheret's SeleniumBasic users who need a way to update the WebDriver in C:\Users\username\AppData\Local\SeleniumBasic
    'there may be a permission issue for writing to this directory so you may have to run as administrator
    Dim mngr As SeleniumVBA.WebDriverManager
    
    Set mngr = SeleniumVBA.New_WebDriverManager
    
    mngr.DefaultDriverFolder = mngr.GetSeleniumBasicFolderPath
    
    MsgBox mngr.AlignEdgeDriverWithBrowser("edgedriver.exe")
    MsgBox mngr.AlignChromeDriverWithBrowser("chromedriver.exe")
    MsgBox mngr.AlignFirefoxDriverWithBrowser("geckodriver.exe")
End Sub
