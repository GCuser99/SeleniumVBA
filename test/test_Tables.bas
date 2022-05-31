Attribute VB_Name = "test_Tables"

Sub test_table()
    'see https://www.guru99.com/selenium-webtable.html
    Dim driver As New WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    'how to write XPath for table in Selenium
    driver.NavigateTo "https://demo.guru99.com/test/write-xpath-table.html"
    driver.Wait 2000
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[1]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[1]/td[2]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[2]").GetText
    
    'accessing nested tables
    driver.NavigateTo "https://demo.guru99.com/test/accessing-nested-table.html"
    driver.Wait 2000
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]").GetText
    
    'using attributes as predicates
    driver.NavigateTo "https://demo.guru99.com/test/newtours/"
    driver.Wait 2000
    Debug.Print driver.FindElement(by.XPath, "//table[@width='270']/tbody/tr[4]/td").GetText
    
    'use inspect element
    Debug.Print driver.FindElement(by.XPath, "//table/tbody/tr/td[2]//table/tbody/tr[4]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td[1]/table[2]/tbody/tr[3]/td[2]/font").GetText
       
    'now see https://www.guru99.com/handling-dynamic-selenium-webdriver.html
    
    Dim webCols As WebElements, webRows As WebElements
    Dim baseTable As WebElement
    Dim tableRow As WebElement, cellIneed As WebElement
    
    'example: fetch number of rows and columns from Dynamic WebTable
    driver.NavigateTo "https://demo.guru99.com/test/web-table-element.php"
    driver.Wait 2000
    
    Set webCols = driver.FindElements(by.XPath, ".//*[@id='leftcontainer']/table/thead/tr/th")
    Debug.Print "No of cols are : " & webCols.Count
    Set webRows = driver.FindElements(by.XPath, ".//*[@id='leftcontainer']/table/tbody/tr/td[1]")
    Debug.Print "No of rows are : " & webRows.Count
    
    'example: fetch cell value of a particular row and column of the Dynamic Table
    Set baseTable = driver.FindElement(by.tagName, "table")
          
    'To find third row of table
    Set tableRow = baseTable.FindElement(by.XPath, "//*[@id='leftcontainer']/table/tbody/tr[3]")
    Debug.Print "Third row of table : " & tableRow.GetText
        
    'to get 3rd row's 2nd column data
    Set cellIneed = tableRow.FindElement(by.XPath, "//*[@id='leftcontainer']/table/tbody/tr[3]/td[2]")
    Debug.Print "Cell value is : " & cellIneed.GetText
    
    'example: get all the values of a Dynamic Table
    driver.NavigateTo "https://demo.guru99.com/test/table.html"
    driver.Wait 2000
    
    Dim mytable As WebElement
    Dim rowsTable As WebElements, columnsRow As WebElements
    Dim row As Integer, col As Integer
    Dim rowElem As WebElement, colElem As WebElement
    
    Set mytable = driver.FindElement(by.XPath, "/html/body/table/tbody")
    Set rowsTable = mytable.FindElements(by.tagName, "tr")
    
    For row = 1 To rowsTable.Count
        Set rowElem = rowsTable(row)
        Set columnsRow = rowElem.FindElements(by.tagName, "td")
        Debug.Print "Number of cells In Row " & row & " are " & columnsRow.Count
        For col = 1 To columnsRow.Count
            Set colElem = columnsRow(col)
            Debug.Print "Cell Value of row number " & row & " and column number " & col & " Is " & colElem.GetText
        Next col
    Next row
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
