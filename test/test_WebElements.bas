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
    
    Debug.Print "Number of rows in table: " & rowsTable.Count
    Debug.Print "Item 1 of first row: " & rowsTable.Item(1).GetText
    Debug.Print "Item 1 of first row: " & rowsTable(1).GetText 'Item is the default property of WebElements class
    Debug.Print "Is member: " & rowsTable.Exists(rowsTable(1))
        
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
    
    Debug.Print "Is member: " & rowsTable.Exists(rowsTable(1))
    
    'can remove a WebElement from collection by index or WebElement object
    rowsTable.Remove 1 'remove first object in collection
    rowsTable.Remove rowsTable(3) 'remove by WebElement object
    Debug.Print "Cells left after removal: " & rowsTable.Count 'after removing two WebElement objects
    
    'works with ExecuteScript when returning a collection of WebElement objects
    Set tableCells = driver.ExecuteScript("return document.getElementsByTagName('td')")
    Debug.Print "Number of table cells: " & tableCells.Count
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
