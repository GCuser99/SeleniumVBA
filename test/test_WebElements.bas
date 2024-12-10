Attribute VB_Name = "test_WebElements"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_WebElements()
    Dim driver As SeleniumVBA.WebDriver
    Dim table As SeleniumVBA.WebElement
    Dim row As SeleniumVBA.WebElement
    Dim rows As SeleniumVBA.WebElements
    Dim cell43_5 As SeleniumVBA.WebElement
    Dim cells As SeleniumVBA.WebElements
    Dim column5 As SeleniumVBA.WebElements
    Dim i As Long
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    driver.NavigateTo "https://the-internet.herokuapp.com/large"
    
    Set table = driver.FindElement(By.ID, "large-table")
    Set rows = table.FindElements(By.TagName, "tr")
    
    'test remove by index, count, item, and exists
    Debug.Assert rows.Count = 51
    rows.Remove 1 'remove header row
    Debug.Assert rows.Count = 50
    Debug.Assert rows.Exists(rows(1)) = True
    
    Set column5 = SeleniumVBA.New_WebElements
    
    'test for each and add
    For Each row In rows
        Set cells = row.FindElements(By.TagName, "td")
        column5.Add cells(5)
    Next row
    
    Set cell43_5 = column5(43)
    
    Debug.Assert cell43_5.GetText = "43.5"
    
    'test remove by object
    column5.Remove cell43_5
    
    Debug.Assert column5(43).GetText = "44.5"

    driver.CloseBrowser
    driver.Shutdown
End Sub
