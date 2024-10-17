Attribute VB_Name = "test_Cookies"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_cookies()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As SeleniumVBA.WebDriver, cks As SeleniumVBA.WebCookies
    
    Set driver = SeleniumVBA.New_WebDriver
    
    Set cks = driver.CreateCookies
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(By.Name, "username").SendKeys ("abc123")
    driver.FindElement(By.Name, "password").SendKeys ("123xyz")
    driver.FindElement(By.Name, "submit").Click
    
    driver.Wait 500
    
    'get all cookies for this domain and then save to file
    driver.GetAllCookies().SaveToFile "cookies.txt"
    
    driver.DeleteAllCookies
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    driver.Wait 1000
    
    'load and set saved cookies from file
    driver.SetCookies cks.LoadFromFile("cookies.txt")
    
    driver.Wait 1000
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cookies2()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As SeleniumVBA.WebDriver, cks As SeleniumVBA.WebCookies, ck As SeleniumVBA.WebCookie

    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(By.Name, "username").SendKeys ("abc123")
    driver.FindElement(By.Name, "password").SendKeys ("123xyz")
    driver.FindElement(By.Name, "submit").Click
    
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
    For Each ck In cks
        If ck.Name = "Selenium" Then
            driver.SetCookie ck
            Exit For
        End If
    Next ck
    
    driver.Wait 500

    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_cookie.php"
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_cookies3()
    'https://www.guru99.com/handling-cookies-selenium-webdriver.html
    Dim driver As SeleniumVBA.WebDriver, cks As SeleniumVBA.WebCookies, ck As SeleniumVBA.WebCookie

    Set driver = SeleniumVBA.New_WebDriver
    
    Set cks = driver.CreateCookies
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/cookie/selenium_aut.php"
    
    driver.Wait 500
    
    driver.FindElement(By.Name, "username").SendKeys ("abc123")
    driver.FindElement(By.Name, "password").SendKeys ("123xyz")
    driver.FindElement(By.Name, "submit").Click
    
    driver.Wait 500
    
    'get cookie add add it to Cookies object
    cks.Add driver.GetCookie("Selenium")
    
    For Each ck In cks
        Debug.Assert ck.Name = "Selenium"
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
