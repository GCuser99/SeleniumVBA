Attribute VB_Name = "test_Cookies"

Sub test_cookies()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver, cks As New Cookies

    Driver.StartChrome

    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    'get all cookies for this domain and then save to file
    Driver.GetAllCookies().SaveToFile ".\cookies.txt"
    
    Driver.DeleteAllCookies
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'load and set saved cookies from file
    Driver.SetCookies cks.LoadFromFile(".\cookies.txt")

    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_cookies2()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver, cks As Cookies, ck As Cookie

    Driver.StartChrome

    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    'get and save the important cookie for this domain
    Set cks = Driver.GetAllCookies
    
    For Each ck In cks
        Debug.Print ck.name
    Next ck
    
    Driver.DeleteAllCookies  'this does not affect the cks object
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'set a specific saved cookie
    Driver.SetCookie cks("Selenium")

    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub

Sub test_cookies3()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim Driver As New WebDriver, cks As New Cookies, ck As Cookie
    
    Driver.StartChrome

    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    Driver.FindElement(by.name, "username").SendKeys ("abc123")
    Driver.FindElement(by.name, "password").SendKeys ("123xyz")
    Driver.FindElement(by.name, "submit").Click
    
    Driver.Wait 500
    
    'get cookie add add it to Cookies object
    cks.Add Driver.GetCookie("Selenium")
    
    For Each ck In cks
        Debug.Print ck.name
    Next ck
    
    'save cookie(s) to file
    cks.SaveToFile ".\cookies.txt"
    
    Driver.DeleteAllCookies 'this does not affect the cks object
    
    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    Driver.Wait 500
    
    'load cookie(s) from file
    cks.LoadFromFile ".\cookies.txt"
    
    'set saved cookie(s) from file
    Driver.SetCookies cks

    Driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    Driver.Wait 500
    Driver.CloseBrowser
    Driver.Shutdown

End Sub
