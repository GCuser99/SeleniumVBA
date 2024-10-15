Attribute VB_Name = "test_WebElements"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_WebElements()
    Dim driver As SeleniumVBA.WebDriver
    Dim myTable As SeleniumVBA.WebElement
    Dim rowsTable As SeleniumVBA.WebElements
    Dim columnsRow As SeleniumVBA.WebElements
    Dim rowElem As SeleniumVBA.WebElement
    Dim colElem As SeleniumVBA.WebElement
    Dim tableCells As SeleniumVBA.WebElements
    Dim row As Integer, col As Integer
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    driver.NavigateTo "https://demo.guru99.com/test/table.html"
    driver.Wait 2000
        
    Set myTable = driver.FindElement(By.XPath, "/html/body/table/tbody")
    Set rowsTable = myTable.FindElements(By.TagName, "tr")
    
    Debug.Assert rowsTable.Count = 5
    Debug.Assert rowsTable.Item(1).GetText = "1 2 3"
    Debug.Assert rowsTable(1).GetText = "1 2 3" 'Item is the default property of WebElements class
    Debug.Assert rowsTable.Exists(rowsTable(1)) = True
        
    'can use the default Item property to iterate through the WebElements object
    For row = 1 To rowsTable.Count
        Set rowElem = rowsTable(row)
        Set columnsRow = rowElem.FindElements(By.TagName, "td")
        For col = 1 To columnsRow.Count
            Set colElem = columnsRow(col)
            Debug.Print "Cell value of row number " & row & " and column number " & col & " is " & colElem.GetText
        Next col
    Next row
    
    'can also use For Each syntax to do same ...
    For Each rowElem In rowsTable
        Set columnsRow = rowElem.FindElements(By.TagName, "td")
        For Each colElem In columnsRow
            Debug.Print "Cell Value is " & colElem.GetText
        Next colElem
    Next rowElem
    
    Debug.Assert rowsTable.Exists(rowsTable(1)) = True
    
    'can remove a WebElement from collection by index or WebElement object
    rowsTable.Remove 1 'remove first object in collection
    rowsTable.Remove rowsTable(3) 'remove by WebElement object
    Debug.Assert rowsTable.Count = 3 'after removing two WebElement objects
    
    'works with ExecuteScript when returning a collection of WebElement objects
    Set tableCells = driver.ExecuteScript("return document.getElementsByTagName('td')")
    Debug.Assert tableCells.Count = 12
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
