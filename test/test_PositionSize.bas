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

    driver.NavigateTo "https://www.wikipedia.org/"
    Set webElem = driver.FindElement(By.ID, "searchInput")

    driver.Wait 500
    
    'SeleniumVBA uses the dictionary object to represent rectangle position and size
    Set rect = webElem.GetRect
    
    Debug.Print "element position/size", rect("x"), rect("y"), rect("width"), rect("height")

    driver.CloseBrowser
    driver.Shutdown
End Sub
