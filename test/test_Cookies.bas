Attribute VB_Name = "test_Cookies"
Sub test_cookies()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver, cks As New WebCookies
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(by.Name, "username").SendKeys ("abc123")
    driver.FindElement(by.Name, "password").SendKeys ("123xyz")
    driver.FindElement(by.Name, "submit").Click
    
    driver.Wait 500
    
    'get all cookies for this domain and then save to file
    driver.GetAllCookies().SaveToFile ".\cookies.txt"
    
    driver.DeleteAllCookies
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 1000
    
    'load and set saved cookies from file
    driver.SetCookies cks.LoadFromFile(".\cookies.txt")
    
    driver.Wait 1000
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cookies2()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver, cks As WebCookies, ck As WebCookie

    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(by.Name, "username").SendKeys ("abc123")
    driver.FindElement(by.Name, "password").SendKeys ("123xyz")
    driver.FindElement(by.Name, "submit").Click
    
    'get and save the important cookie for this domain
    Set cks = driver.GetAllCookies
    
    For Each ck In cks
        Debug.Print ck.Name
    Next ck
    
    driver.DeleteAllCookies  'this does not affect the cks object
    
    driver.Wait 500
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 500
    
    'set a specific saved cookie
    driver.SetCookie cks("Selenium")
    
    driver.Wait 500

    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cookies3()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver, cks As New WebCookies, ck As WebCookie

    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(by.Name, "username").SendKeys ("abc123")
    driver.FindElement(by.Name, "password").SendKeys ("123xyz")
    driver.FindElement(by.Name, "submit").Click
    
    driver.Wait 500
    
    'get cookie add add it to Cookies object
    cks.Add driver.GetCookie("Selenium")
    
    For Each ck In cks
        Debug.Print ck.Name
    Next ck
    
    'save cookie(s) to file
    cks.SaveToFile ".\cookies.txt"
    
    driver.DeleteAllCookies 'this does not affect the cks object
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 500
    
    'load cookie(s) from file
    cks.LoadFromFile ".\cookies.txt"
    
    'set saved cookie(s) from file
    driver.SetCookies cks
    
    driver.Wait 500

    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
