Attribute VB_Name = "test_Extensions"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'Note: To add extensions to Chrome and Edge browsers, you must use the BiDi add-on (see https://github.com/hanamichi77777/WebDriverBiDi-via-VBA-test)

Sub test_InstallAddon()
    Dim driver As SeleniumVBA.WebDriver
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartFirefox
    driver.OpenBrowser
    
    'this will install an xpi add-on (Firefox only) - use AddExtensions method of WebCapabilities for Edge/Chrome
    driver.InstallAddon Environ("USERPROFILE") & "\Documents\SeleniumVBA\extensions\darkreader-4.9.94.xpi"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
