Attribute VB_Name = "test_PositionSize"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_position_size()
    Dim driver As SeleniumVBA.WebDriver
    Dim webElem As SeleniumVBA.WebElement
    Dim rect As Dictionary

    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser

    driver.NavigateTo "https://www.selenium.dev/selenium/web/single_text_input.html"
    Set webElem = driver.FindElement(By.ID, "textInput")

    driver.Wait 500
    
    'SeleniumVBA uses the dictionary object to represent rectangle position and size
    Set rect = webElem.GetRect
    
    Debug.Assert rect("x") = 8
    Debug.Assert rect("y") = 8
    Debug.Assert rect("width") = 173
    Debug.Assert rect("height") = 21
    Debug.Assert rect("left") = 8
    Debug.Assert rect("top") = 8
    Debug.Assert rect("bottom") = 29
    Debug.Assert rect("right") = 181

    driver.CloseBrowser
    driver.Shutdown
End Sub
