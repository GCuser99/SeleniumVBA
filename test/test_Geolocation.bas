Attribute VB_Name = "test_Geolocation"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_geolocation()
    'if running in incognito mode, then consider setting SetGeolocationAware
    'method of WebCapabilities object to True (see test_Capabilities module)
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome 'Chrome and Edge only
    driver.OpenBrowser
    
    'set the location
    driver.SetGeolocation 41.1621429, -8.6219537
  
    driver.NavigateTo "https://whatmylocation.com/"
    driver.Wait 2000
    
    'print the name/address of the location to immediate window
    Debug.Print driver.FindElementByXPath("//*[@id='address']").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
