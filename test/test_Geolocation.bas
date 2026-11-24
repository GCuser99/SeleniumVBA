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
    
    'set the location and grant geolocation awareness
    driver.SetGeolocation latitude:=41.1621429, longitude:=-8.6219537
    driver.AllowGeolocation
    
    driver.NavigateTo "https://the-internet.herokuapp.com/geolocation"
    
    driver.FindElementByXPath("//*[@id='content']/div/button").Click
    
    Debug.Assert driver.FindElementByID("lat-value").GetText = 41.1621429
    Debug.Assert driver.FindElementByID("long-value").GetText = -8.6219537
    
    Debug.Assert driver.GetGeolocation()("latitude") = 41.1621429
    Debug.Assert driver.GetGeolocation()("longitude") = -8.6219537
    
    driver.FindElementByXPath("//*[@id='map-link']/a").Click
    
    driver.Wait 3000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
