Attribute VB_Name = "test_Tables"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_table()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    html = "<html><head><title>Test Table To Array</title></head><body><table border='l' id='mytable'>"
    html = html & "<thead><tr><th>head 1</th><th>head 2</th></tr></thead>"
    html = html & "<tbody><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>"
    html = html & "<table border='l'><tbody><tr><td>4A</td><td>4B</td></tr><tr><td>4C</td><td>4D</td></tr></tbody></table>"
    html = html & "</td></tr></tbody><tfoot><tr><td colspan='2'>footer content</td></tr></tfoot></table></body></html>"
    
    driver.NavigateToString html
    
    driver.Wait 1000
    
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/thead/tr[1]/th[1]").GetText = "head 1"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/thead/tr[1]/th[2]").GetText = "head 2"
    
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[1]").GetText = "1"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[2]").GetText = "2"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[1]").GetText = "3"
    
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[1]").GetText = "4A"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]").GetText = "4B"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[1]").GetText = "4C"
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]").GetText = "4D"
    
    Debug.Assert driver.FindElement(By.XPath, "//table[@id='mytable']/tfoot/tr[1]/td[1]").GetText = "footer content"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_table_to_array()
    Dim driver As SeleniumVBA.WebDriver
    Dim table() As Variant
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    html = "<html><head><title>Test Table To Array</title></head><body><table border='l' id='mytable'><thead><tr><th>head 1</th><th>head 2</th><th>head 3</th></tr></thead><tbody><tr><td>Dos Equis:</td><td colspan='2'>X</td></tr><tr><td>Choose</td><td>Option</td><td><table border='l'><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table></td></tr><tr><td>Sky</td><td rowspan='3'>Is</td><td>Blue</td></tr><tr><td>Less</td><td>More</td></tr><tr><td>Big</td><td rowspan='2'>Better</td></tr><tr><td>I</td><td>Feel</td></tr></tbody><tfoot><tr><td colspan='3'>footer content</td></tr></tfoot></table></body></html>"
    
    driver.NavigateToString html
    
    driver.Wait 1000
    
    table = driver.FindElement(By.ID, "mytable").TableToArray()
    
    'With createSpanData=True (default):
    Debug.Assert table(1, 1) & " " & table(1, 2) & " " & table(1, 3) = "head 1 head 2 head 3"
    Debug.Assert table(2, 1) & " " & table(2, 2) & " " & table(2, 3) = "Dos Equis: X X"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(1, 1) = "Choose Option A"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(1, 2) = "Choose Option B"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(2, 1) = "Choose Option C"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(2, 2) = "Choose Option D"
    Debug.Assert table(4, 1) & " " & table(4, 2) & " " & table(4, 3) = "Sky Is Blue"
    Debug.Assert table(5, 1) & " " & table(5, 2) & " " & table(5, 3) = "Less Is More"
    Debug.Assert table(6, 1) & " " & table(6, 2) & " " & table(6, 3) = "Big Is Better"
    Debug.Assert table(7, 1) & " " & table(7, 2) & " " & table(7, 3) = "I Feel Better"
    Debug.Assert table(8, 1) & " " & table(8, 2) & " " & table(8, 3) = "footer content footer content footer content"
    
    'now process table w/o creating span data
    
    table = driver.FindElement(By.ID, "mytable").TableToArray(createSpanData:=False)
    
    'With createSpanData=False:
    Debug.Assert table(1, 1) & " " & table(1, 2) & " " & table(1, 3) = "head 1 head 2 head 3"
    Debug.Assert table(2, 1) & " " & table(2, 2) & " " & table(2, 3) = "Dos Equis: X "
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(1, 1) = "Choose Option A"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(1, 2) = "Choose Option B"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(2, 1) = "Choose Option C"
    Debug.Assert table(3, 1) & " " & table(3, 2) & " " & table(3, 3)(2, 2) = "Choose Option D"
    Debug.Assert table(4, 1) & " " & table(4, 2) & " " & table(4, 3) = "Sky Is Blue"
    Debug.Assert table(5, 1) & " " & table(5, 2) & " " & table(5, 3) = "Less More "
    Debug.Assert table(6, 1) & " " & table(6, 2) & " " & table(6, 3) = "Big Better "
    Debug.Assert table(7, 1) & " " & table(7, 2) & " " & table(7, 3) = "I Feel "
    Debug.Assert table(8, 1) & " " & table(8, 2) & " " & table(8, 3) = "footer content  "

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_large_table_to_array()
    Dim driver As SeleniumVBA.WebDriver
    Dim table() As Variant
    Dim html As String
    Dim i As Long, j As Long
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    'build the large table page (in this case having 20,000 cells)
    html = "<html><head><title>Test Table To Array</title></head><body><table border='1' id='mytable'>"
    For i = 1 To 200
        html = html & "<tr>"
        For j = 1 To 100
            html = html & "<td>" & i & "." & j & "</td>"
        Next j
        html = html & "</tr>"
    Next i
    html = html & "</table></body></html>"
    
    driver.NavigateToString html
    
    'this is super-fast due to generating the table in-browser using JavaScript
    table = driver.FindElement(By.ID, "mytable").TableToArray(createSpanData:=False)
    
    Debug.Assert table(7, 30) = "7.30"
    Debug.Assert table(99, 2) = "99.2"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_table_to_array_formatting()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim table() As Variant
    Dim html As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    html = "<html><head><title>Test Table To Array Formatting</title></head><body>"
    html = html & "<table border='l' id='mytable'><tr>"
    html = html & "<td>12/14/2024<br>12/15/2024</td>"
    html = html & "<td>Hi,&nbsp;this&nbsp;is&nbsp;<p>Mike</p></td>"
    html = html & "</tr></table></body></html>"
    driver.NavigateToString html
    
    driver.Wait 1500
    
    Set elem = driver.FindElementByCssSelector("#mytable")
    
    table = elem.TableToArray(ignoreCellFormatting:=False) 'default setting
    
    Debug.Assert table(1, 1) = "12/14/2024" & vbCrLf & "12/15/2024"
    Debug.Assert table(1, 2) = "Hi, this is " & vbCrLf & vbCrLf & "Mike"

    table = elem.TableToArray(ignoreCellFormatting:=True)
    
    Debug.Assert table(1, 1) = "12/14/202412/15/2024"
    Debug.Assert table(1, 2) = "Hi, this is Mike"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
