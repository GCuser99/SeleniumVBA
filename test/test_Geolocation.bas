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
    
    driver.ImplicitMaxWait = 2000
    
    'set the location
    driver.SetGeolocation 41.1621429, -8.6219537
  
    driver.NavigateTo "https://the-internet.herokuapp.com/geolocation"
    
    driver.FindElementByXPath("//*[@id='content']/div/button").Click
    
    Debug.Print driver.FindElementByID("lat-value").GetText, driver.FindElementByID("long-value").GetText
    
    driver.Wait 2000
    
    driver.FindElementByXPath("//*[@id='map-link']/a").Click
    
    driver.Wait 5000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

