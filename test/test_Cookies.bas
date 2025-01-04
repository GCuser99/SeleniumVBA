Attribute VB_Name = "test_Cookies"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_session_cookie()
    Dim driver As SeleniumVBA.WebDriver
    Dim cks As SeleniumVBA.WebCookies
    Dim ck As SeleniumVBA.WebCookie
    
    Set driver = SeleniumVBA.New_WebDriver
    
    Set cks = driver.CreateCookies

    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/sessionCookie.html"
    driver.Wait 500
    
    'this creates a cookie named bgcolor containing the background color as its value
    driver.FindElement(By.CssSelector, "#setcolorbutton").Click
    driver.Wait 500
    
    'get cookie for this domain and then save to file
    driver.GetAllCookies().SaveToFile "cookies.txt"
    
    'click to open the new window - this tries to set the background color through the passed cookie
    'but note that because the cookie was deleted, the background color does not get set
    driver.FindElement(By.CssSelector, "#openwindowbutton").Click
    driver.Wait 500
    
    driver.Windows.SwitchToByTitle "Session cookie destination*"
    Debug.Assert driver.ExecuteScript("return document.body.style.backgroundColor;") = "rgb(128, 255, 255)"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/sessionCookie.html"
    
    'make sure all cookies are deleted from the session
    driver.DeleteAllCookies
    
    'load the cookies from previous session
    driver.SetCookies cks.LoadFromFile("cookies.txt")
    
    For Each ck In cks
        Debug.Assert ck.Name = "bgcolor"
        Debug.Assert ck.Value = "#80FFFF"
    Next ck
    
    'click to open the new window - this sets the background color through the loaded cookie
    driver.FindElement(By.CssSelector, "#openwindowbutton").Click
    driver.Wait 500
    
    driver.Windows.SwitchToByTitle "Session cookie destination*"
    Debug.Assert driver.ExecuteScript("return document.body.style.backgroundColor;") = "rgb(128, 255, 255)"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/sessionCookie.html"
    
    'make sure all cookies are deleted from the session
    driver.DeleteAllCookies
    
    'change the value (background color) of the cookie from from previous session and set it
    cks(1).Value = "#8080ff"
    driver.SetCookie cks(1)
    
    'click to open the new window - this sets the background color through the loaded cookie
    driver.FindElement(By.CssSelector, "#openwindowbutton").Click
    driver.Wait 500
    
    driver.Windows.SwitchToByTitle "Session cookie destination*"
    Debug.Assert driver.ExecuteScript("return document.body.style.backgroundColor;") = "rgb(128, 128, 255)"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
