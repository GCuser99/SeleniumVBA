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
        
    html = "<html><head><title>Test Table To Array</title></head><body>"
        
    html = html & "<table border='l' id='mytable0'><thead><tr><th>head 1</th><th>head 2</th><th>head 3</th></tr></thead><tbody><tr><td>Dos Equis:</td><td colspan='2'>X</td></tr><tr><td>Choose</td><td>Option</td><td><table border='l'><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table></td></tr><tr><td>Sky</td><td rowspan='3'>Is</td><td>Blue</td></tr><tr><td>Less</td><td>More</td></tr><tr><td>Big</td><td rowspan='2'>Better</td></tr><tr><td>I</td><td>Feel</td></tr></tbody><tfoot><tr><td colspan='3'>footer content</td></tr></tfoot></table>"
        
    html = html & "<div style='height:20px;'></div>"
    
    html = html & "<table border='1' id='mytable1'>"
    html = html & "<thead><tr><td>H1</td></tr><tr><td>H2</td></tr><tr><td>H3</td></tr></thead>"
    html = html & "<tbody><tr><td>D1</td></tr><tr><td>D2</td></tr></tbody>"
    html = html & "<tfoot><tr><td>F1</td></tr><tr><td>F2</td></tr></tfoot></table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable2'>"
    html = html & "<tr><td>R1</td></tr>"
    html = html & "<tr><td><table border='1'><thead><tr><td>x</td></tr></thead><tbody><tr><td>y</td></tr></tbody></table></td></tr>"
    html = html & "</table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable3'>"
    html = html & "<tr><td rowspan='2' colspan='2'>A</td><td>B</td></tr>"
    html = html & "<tr><td>C</td></tr>"
    html = html & "<tr><td>D</td><td colspan='2'>E</td></tr>"
    html = html & "<tr><td>F</td><td>G</td><td>H</td></tr>"
    html = html & "</table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable4'>"
    html = html & "<thead><tr><td>H1</td></tr></thead><tbody></tbody>"
    html = html & "</table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable5'><tr><td rowspan='3' colspan='3'>A</td><td>B</td></tr><tr><td>C</td></tr><tr><td>D</td></tr><tr><td>1</td><td>2</td><td>3</td><td>4</td></tr></table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable6'><tr><td rowspan='5'>A</td><td>1</td></tr><tr><td>2</td></tr></table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable7'><tr><td rowspan='2'>A</td><td colspan='2'>B</td></tr><tr><td>C</td><td>D</td></tr></table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable8'><tr><td rowspan='2'>A</td><td colspan='3'>B</td></tr><tr><td>C</td><td>D</td><td>E</td></tr></table>"
        
    html = html & "<div style='height:20px;'></div>"
        
    html = html & "<table border='1' id='mytable9'><tr><td rowspan='2'>A</td><td>B</td><td>X</td></tr><tr><td colspan='2'>C</td></tr></table>"
        
    html = html & "</body></html>"
        
    driver.NavigateToString html
    
    driver.Wait 1000
        
    table = driver.FindElement(By.ID, "mytable0").TableToArray()
    
    ' With createSpanData=True (default):
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
    
    table = driver.FindElement(By.ID, "mytable0").TableToArray(createSpanData:=False)
    
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
        
    'no skipping - all 7 rows
    table = driver.FindElement(By.ID, "mytable1").TableToArray()
    Debug.Assert UBound(table, 1) = 7
    
    'skip all 3 header rows
    table = driver.FindElement(By.ID, "mytable1").TableToArray(skipHeader:=True)
    Debug.Assert UBound(table, 1) = 4
    Debug.Assert table(1, 1) = "D1"
    
    'skip all 2 footer rows
    table = driver.FindElement(By.ID, "mytable1").TableToArray(skipFooter:=True)
    Debug.Assert UBound(table, 1) = 5
    Debug.Assert table(5, 1) = "D2"
    
    'skip both
    table = driver.FindElement(By.ID, "mytable1").TableToArray(skipHeader:=True, skipFooter:=True)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert table(1, 1) & " " & table(2, 1) = "D1 D2"
        
    'the outer table has no thead - only the nested table does
    'outer rows are kept, nested header is removed
    table = driver.FindElement(By.ID, "mytable2").TableToArray(skipHeader:=True)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert table(1, 1) = "R1"
    Debug.Assert UBound(table(2, 1), 1) = 1
    Debug.Assert table(2, 1)(1, 1) = "y"
    
    'A spans 2 rows and 2 columns
    table = driver.FindElement(By.ID, "mytable3").TableToArray()
    Debug.Assert UBound(table, 1) = 4
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "AAB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "AAC"
    Debug.Assert table(3, 1) & table(3, 2) & table(3, 3) = "DEE"
    Debug.Assert table(4, 1) & table(4, 2) & table(4, 3) = "FGH"
    table = driver.FindElement(By.ID, "mytable3").TableToArray(createSpanData:=False)
    Debug.Assert UBound(table, 1) = 4
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "AB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "C"
    Debug.Assert table(3, 1) & table(3, 2) & table(3, 3) = "DE"
    Debug.Assert table(4, 1) & table(4, 2) & table(4, 3) = "FGH"
        
    Dim errNum As Long, errDesc As String
    On Error Resume Next
    table = driver.FindElement(By.ID, "mytable4").TableToArray(skipHeader:=True)
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    Debug.Assert errNum = 1
    Debug.Assert InStr(errDesc, "the table has no data rows") > 0
        
    'a cell spanning 3 rows and 3 columns
    table = driver.FindElement(By.ID, "mytable5").TableToArray()
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) & table(1, 4) = "AAAB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) & table(2, 4) = "AAAC"
    Debug.Assert table(3, 1) & table(3, 2) & table(3, 3) & table(3, 4) = "AAAD"
    Debug.Assert table(4, 1) & table(4, 2) & table(4, 3) & table(4, 4) = "1234"
    table = driver.FindElement(By.ID, "mytable5").TableToArray(createSpanData:=False)
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) & table(1, 4) = "AB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) & table(2, 4) = "C"
    Debug.Assert table(3, 1) & table(3, 2) & table(3, 3) & table(3, 4) = "D"
    Debug.Assert table(4, 1) & table(4, 2) & table(4, 3) & table(4, 4) = "1234"
        
    'rowspan larger than the remaining rows
    table = driver.FindElement(By.ID, "mytable6").TableToArray()
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert table(1, 1) & table(1, 2) & " " & table(2, 1) & table(2, 2) = "A1 A2"
    'rowspan larger than the remaining rows
    table = driver.FindElement(By.ID, "mytable6").TableToArray(createSpanData:=False)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert table(1, 1) & table(1, 2) & " " & table(2, 1) & table(2, 2) = "A1 2"
        
    ' A spans 2 rows, B spans 2 columns
    table = driver.FindElement(By.ID, "mytable7").TableToArray()
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "ABB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "ACD"
    table = driver.FindElement(By.ID, "mytable7").TableToArray(createSpanData:=False)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 2
    Debug.Assert table(1, 1) & table(1, 2) = "AB"
    Debug.Assert table(2, 1) & table(2, 2) = "CD"
        
    ' B spans 3 columns
    table = driver.FindElement(By.ID, "mytable8").TableToArray()
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 4
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) & table(1, 4) = "ABBB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) & table(2, 4) = "ACDE"
    table = driver.FindElement(By.ID, "mytable8").TableToArray(createSpanData:=False)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "AB"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "CDE"
        
    ' the column span is in the row BELOW the row span
    table = driver.FindElement(By.ID, "mytable9").TableToArray()
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "ABX"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "ACC"
    table = driver.FindElement(By.ID, "mytable9").TableToArray(createSpanData:=False)
    Debug.Assert UBound(table, 1) = 2
    Debug.Assert UBound(table, 2) = 3
    Debug.Assert table(1, 1) & table(1, 2) & table(1, 3) = "ABX"
    Debug.Assert table(2, 1) & table(2, 2) & table(2, 3) = "C"

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_large_table_to_array()
    Dim driver As WebDriver
    Dim table() As Variant
    Dim html As String
    Dim i As Long, j As Long
    
    Set driver = New WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    'build the large table page (in this case having 20,000 cells)
    html = "<html><head><title>Test Table To Array</title></head><body><table border='1' id='mytable'><tbody>"
    For i = 1 To 200
        html = html & "<tr>"
        For j = 1 To 100
            html = html & "<td>" & i & "." & j & "</td>"
        Next j
        html = html & "</tr>"
    Next i
    html = html & "</tbody></table></body></html>"
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

