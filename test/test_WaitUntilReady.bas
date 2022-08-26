Attribute VB_Name = "test_WaitUntilReady"
Option Explicit
Option Private Module

Sub test_WaitUntilReady()
    Dim driver As SeleniumVBA.WebDriver
    Dim searchButton As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    
    Set searchButton = driver.FindElement(by.Name, "btnK")
    
    driver.Wait 500
    
    'search button is there, but not interactable...
    Debug.Print "Is search button interactable yet? " & searchButton.IsDisplayed
    
    driver.FindElement(by.Name, "q").SendKeys "Interactable"
    
    'searchButton.Click 'will often throw an error here because it takes some time
    'for search button to get ready after typing search phrase
    Debug.Print "Is search button interactable yet? " & searchButton.IsDisplayed
    
    'can place an explicit Wait here but another way is to use WaitUntilReady method
    'it returns the "ready" input element object so can use methods on same line
    searchButton.WaitUntilReady().Click
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

