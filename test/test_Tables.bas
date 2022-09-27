Attribute VB_Name = "test_Tables"
Option Explicit
Option Private Module

Sub test_table()
    Dim driver As SeleniumVBA.WebDriver
    Dim v() As Variant, htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    htmlStr = "<html><body><table border='l' id='mytable'><thead><tr><th>head 1</th><th>head 2</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr><tr><td>3</td><td><table border='l'><tbody><tr><td>4A</td><td>4B</td></tr><tr><td>4C</td><td>4D</td></tr></tbody></table></td></tr></tbody><tfoot><tr><td colspan='2'>footer content</td></tr></tfoot></table></body></html>"
    driver.SaveHTMLToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 1000
    
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/thead/tr[1]/th[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/thead/tr[1]/th[2]").GetText
    
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[1]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[1]/td[2]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[2]/td[1]").GetText
    
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[1]").GetText
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]").GetText
    
    Debug.Print driver.FindElement(by.XPath, "//table[@id='mytable']/tfoot/tr[1]/td[1]").GetText
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_table_to_array()
    Dim driver As SeleniumVBA.WebDriver
    Dim v() As Variant, htmlStr As String, i, j, k
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    htmlStr = "<html><body><table border='l' id='mytable'><thead><tr><th>head 1</th><th>head 2</th><th>head 3</th></tr></thead><tbody><tr><td>Dos Equis:</td><td colspan='2'>X</td></tr><tr><td>Choose</td><td>Option</td><td><table border='l'><tbody><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></tbody></table></td></tr><tr><td>Sky</td><td rowspan='3'>Is</td><td>Blue</td></tr><tr><td>Less</td><td>More</td></tr><tr><td>Big</td><td rowspan='2'>Better</td></tr><tr><td>I</td><td>Feel</td></tr></tbody><tfoot><tr><td colspan='3'>footer content</td></tr></tfoot></table></body></html>"
    
    driver.SaveHTMLToFile htmlStr, ".\snippet.html"
    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 5000
    
    v = driver.FindElement(by.ID, "mytable").TableToArray()
    
    Debug.Print "With createSpanData=True (default):"
    For i = 1 To UBound(v, 1)
        If Not IsArray(v(i, 3)) Then
            Debug.Print v(i, 1), v(i, 2), v(i, 3)
        Else
            For j = 1 To UBound(v(i, 3), 1)
                For k = 1 To UBound(v(i, 3), 2)
                    Debug.Print v(i, 1), v(i, 2), v(i, 3)(j, k)
                Next k
            Next j
        End If
    Next i
    
    'now process table w/o creating span data
    
    v = driver.FindElement(by.ID, "mytable").TableToArray(createSpanData:=False)
    
    Debug.Print ""
    Debug.Print "With createSpanData=False:"
    For i = 1 To UBound(v, 1)
        If Not IsArray(v(i, 3)) Then
            Debug.Print v(i, 1), v(i, 2), v(i, 3)
        Else
            For j = 1 To UBound(v(i, 3), 1)
                For k = 1 To UBound(v(i, 3), 2)
                    Debug.Print v(i, 1), v(i, 2), v(i, 3)(j, k)
                Next k
            Next j
        End If
    Next i

    driver.CloseBrowser
    driver.Shutdown
End Sub


