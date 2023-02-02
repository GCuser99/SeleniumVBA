Attribute VB_Name = "test_UserAgent"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_hide_headless()
    'some servers detect headless mode in the sent User Agent and then deny access, so
    'here we modify the user agent that gets sent to server
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    Dim ua As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser invisible:=True  'a way of running headless mode without explicitly adding --headless arg to Capabilities
    
    'get the user agent for this browser setup
    ua = driver.GetUserAgent
    
    Debug.Print "Original User Agent:  " & ua
    
    driver.CloseBrowser
    
    Set caps = driver.CreateCapabilities
    
    'now we modify the user agent string by tossing the "Headless" keyword and then
    'update WebCapabilities UserArgent argument
    caps.SetUserAgent Replace$(ua, "HeadlessChrome", "Chrome")
    
    driver.OpenBrowser caps, invisible:=True
    
    'to see a full list of headers navigate to https://www.httpbin.org/headers
    driver.NavigateTo "https://www.whatismybrowser.com/detect/what-is-my-user-agent/"
    
    Debug.Print "Modfified User Agent: " & driver.FindElement(By.ID, "detected_value").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
