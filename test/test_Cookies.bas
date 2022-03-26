Attribute VB_Name = "test_Cookies"

Sub test_cookies()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver

    driver.StartChrome

    driver.OpenBrowser
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.FindElement(by.name, "username").SendKeys ("abc123")
    driver.FindElement(by.name, "password").SendKeys ("123xyz")
    driver.FindElement(by.name, "submit").Click
    
    driver.Wait 500
    
    'get all cookies for this domain and then save to file
    driver.SaveCookiesToFile driver.GetAllCookies(), ".\cookies.txt"
    
    driver.DeleteAllCookies
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 500
    
    'load and set saved cookies from file
    driver.SetCookies driver.LoadCookiesFromFile(".\cookies.txt")

    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    driver.CloseBrowser
    driver.Shutdown

End Sub

Sub test_cookie()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver, ck As Cookie

    driver.StartChrome

    driver.OpenBrowser
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.FindElement(by.name, "username").SendKeys ("abc123")
    driver.FindElement(by.name, "password").SendKeys ("123xyz")
    driver.FindElement(by.name, "submit").Click
    
    driver.Wait 500
    
    'get and save the important cookie for this domain
    Set ck = driver.GetCookie("Selenium")
    
    driver.DeleteAllCookies
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 500
    
    'set saved cookie
    driver.SetCookie ck

    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    driver.CloseBrowser
    driver.Shutdown

End Sub

Sub test_cookie2()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As New WebDriver, cks() As Cookie
    
    driver.StartChrome

    driver.OpenBrowser
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.FindElement(by.name, "username").SendKeys ("abc123")
    driver.FindElement(by.name, "password").SendKeys ("123xyz")
    driver.FindElement(by.name, "submit").Click
    
    driver.Wait 500
    
    
    'get cookie
    ReDim cks(0 To 0) 'save cookies to file needs a cookie array
    Set cks(0) = driver.GetCookie("Selenium")
    
    'save cookie to file
    driver.SaveCookiesToFile cks, ".\cookies.txt"
    
    driver.DeleteAllCookies
    
    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 500
    
    'load cookie from file
    Set cks(0) = driver.LoadCookiesFromFile(".\cookies.txt")(0)
    
    'set saved cookies from file
    driver.SetCookie cks(0)

    driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    driver.CloseBrowser
    driver.Shutdown

End Sub
