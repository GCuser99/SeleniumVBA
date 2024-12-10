Attribute VB_Name = "test_Tables"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_table()
    Dim driver As SeleniumVBA.WebDriver
    Dim htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser

    htmlStr = "<html><body><table border='l' id='mytable'><thead><tr><th>head 1</th><th>head 2</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr><tr><td>3</td><td><table border='l'><tbody><tr><td>4A</td><td>4B</td></tr><tr><td>4C</td><td>4D</td></tr></tbody></table></td></tr></tbody><tfoot><tr><td colspan='2'>footer content</td></tr></tfoot></table></body></html>"
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    
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
    
    driver.DeleteFiles ".\snippet.html"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_table_to_array()
    Dim driver As SeleniumVBA.WebDriver
    Dim table() As Variant
    Dim htmlStr As String
    Dim i As Long, j As Long, k As Long
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    htmlStr = "<html><body><table border='l' id='mytable'><thead><tr><th>head 1</th><th>head 2</th><th>head 3</th></tr></thead><tbody><tr><td>Dos Equis:</td><td colspan='2'>X</td></tr><tr><td>Choose</td><td>Option</td><td><table border='l'><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table></td></tr><tr><td>Sky</td><td rowspan='3'>Is</td><td>Blue</td></tr><tr><td>Less</td><td>More</td></tr><tr><td>Big</td><td rowspan='2'>Better</td></tr><tr><td>I</td><td>Feel</td></tr></tbody><tfoot><tr><td colspan='3'>footer content</td></tr></tfoot></table></body></html>"
    
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    driver.NavigateToFile ".\snippet.html"
    
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

Sub test_table_to_array_large()
    Dim driver As SeleniumVBA.WebDriver
    Dim table() As Variant
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 2000
    
    driver.NavigateTo "https://the-internet.herokuapp.com/large"
    
    table = driver.FindElement(By.ID, "large-table").TableToArray(skipHeader:=True)
    
    Debug.Assert UBound(table, 1) = 50
    Debug.Assert UBound(table, 2) = 50
    Debug.Assert table(43, 5) = "43.5"

    driver.CloseBrowser
    driver.Shutdown
End Sub
