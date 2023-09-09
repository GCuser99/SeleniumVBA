Attribute VB_Name = "test_Scroll"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_scroll_ops()
    'see https://www.guru99.com/scroll-up-down-selenium-webdriver.html
    'for more info and tutorials see https://www.guru99.com/selenium-tutorial.html
    Dim driver As SeleniumVBA.WebDriver
    Dim webElem As SeleniumVBA.WebElement
    Dim i As Integer
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "http://demo.guru99.com/test/guru99home/"
    driver.ActiveWindow.Maximize
    driver.Wait 1000
    
    'scroll down in increments of 50 pixels
    For i = 1 To 40
        driver.ScrollBy , 50
        driver.Wait 25
    Next i
    
    'scroll to top of window
    driver.ScrollToTop
    driver.Wait 1000
    
    'scroll down half-way from top to bottom of window
    driver.ScrollTo 0, driver.GetScrollHeight / 2
    driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll vertically to a WebElement
    driver.FindElement(By.LinkText, "Linux").ScrollIntoView
    
    driver.Wait 1000
    
    'This will scroll the webpage to end.
    driver.ScrollToBottom
    driver.Wait 1000
    
    driver.NavigateTo "http://demo.guru99.com/test/guru99home/scrolling.html"
    driver.Wait 1000
    
    'this uses the WebElement ScrollIntoView method to scroll horizontally to a WebElement
    driver.FindElement(By.LinkText, "VBScript").ScrollIntoView
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
