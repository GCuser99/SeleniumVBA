Attribute VB_Name = "test_cookies"

Sub test_cookies()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver

    Driver.Chrome

    Driver.OpenBrowser
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    'get all cookies for this domain and then save to file
    Driver.SaveCookiesToFile Driver.GetAllCookies(), ".\cookies.txt"
    
    Driver.DeleteAllCookies
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'load and set saved cookies from file
    Driver.SetCookies Driver.LoadCookiesFromFile(".\cookies.txt")

    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_cookie()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver, ck As Cookie

    Driver.Chrome

    Driver.OpenBrowser
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    'get and save the important cookie for this domain
    Set ck = Driver.GetCookie("Selenium")
    
    Driver.DeleteAllCookies
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'set saved cookie
    Driver.SetCookie ck

    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_cookie2()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver, cks() As Cookie
    Driver.Chrome

    Driver.OpenBrowser
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    
    'get cookie
    ReDim cks(0 To 0) 'save cookies to file needs a cookie array
    Set cks(0) = Driver.GetCookie("Selenium")
    
    'save cookie to file
    Driver.SaveCookiesToFile cks, ".\cookies.txt"
    
    Driver.DeleteAllCookies
    
    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'load cookie from file
    Set cks(0) = Driver.LoadCookiesFromFile(".\cookies.txt")(0)
    
    'set saved cookies from file
    Driver.SetCookie cks(0)

    Driver.Navigate "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
