Attribute VB_Name = "test_Highlight"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_highlight()
    Dim driver As SeleniumVBA.WebDriver
    Dim v() As Variant, htmlStr As String
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartEdge
    driver.OpenBrowser
    
    htmlStr = "<html><body><table border='l' id='mytable'>" _
    & "<thead><tr><th>head 1</th><th>head 2</th></tr></thead>" _
    & "<tbody><tr><td>1</td><td>2</td></tr>" _
    & "<tr><td>3</td><td><table border='l'><tbody>" _
    & "<tr><td>4A</td><td>4B</td></tr><tr><td>4C</td><td>4D</td></tr></tbody>" _
    & "</table></td></tr></tbody>" _
    & "<tfoot><tr><td colspan='2'>footer content</td></tr></tfoot></table></body></html>"
    
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    
    driver.Wait 1000
    
    'automatically highlight every last found element(s):
    driver.SetHightlightFoundElems True
    
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/thead/tr[1]/th[1]").GetText
    driver.Wait
    
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/thead/tr[1]/th[2]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[1]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[2]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[1]").GetText
    driver.Wait
    
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[1]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[1]/td[2]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[1]").GetText
    driver.Wait
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]").GetText
    driver.Wait
    
    Debug.Print driver.FindElement(By.XPath, "//table[@id='mytable']/tfoot/tr[1]/td[1]").GetText
    driver.Wait
    
    driver.SetHightlightFoundElems False
    
    'highlight specified elements (all arguments optional):
    driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[2]/table").Highlight borderColor:=Magenta
    driver.Wait 1500
    driver.FindElement(By.XPath, "//table[@id='mytable']/thead/tr[1]/th[1]").Highlight borderColor:=Blue, unHighlightLast:=False
    driver.Wait 1500
    driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[1]").Highlight borderColor:=Cyan
    driver.Wait 1500
    driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[1]/td[2]").Highlight borderColor:=Green
    driver.Wait 1500
    driver.FindElement(By.XPath, "//table[@id='mytable']/tbody/tr[2]/td[1]").Highlight borderColor:=Black, backgroundColor:=Green
    driver.Wait 1500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_highlight2()
    Dim driver As SeleniumVBA.WebDriver
    Dim elemsBlue As SeleniumVBA.WebElements
    Dim elemsRed As SeleniumVBA.WebElements
    Dim htmlStr As String, i As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    Set elemsBlue = SeleniumVBA.New_WebElements
    Set elemsRed = SeleniumVBA.New_WebElements

    driver.StartEdge
    driver.OpenBrowser
    
    htmlStr = "<html><body><table border='1' id='mytable'>" _
    & "<tr><td>Red</td><td>Blue</td><td>Red</td><td>Blue</td><td>Red</td></tr>" _
    & "<tr><td>Blue</td><td>Red</td><td>Blue</td><td>Red</td><td>Blue</td></tr>" _
    & "<tr><td>Red</td><td>Blue</td><td>Red</td><td>Blue</td><td>Red</td></tr>" _
    & "<tr><td>Blue</td><td>Red</td><td>Blue</td><td>Red</td><td>Blue</td></tr>" _
    & "<tr><td>Red</td><td>Blue</td><td>Red</td><td>Blue</td><td>Red</td></tr>" _
    & "</table></body></html>"
    
    driver.SaveStringToFile htmlStr, ".\snippet.html"
    
    driver.NavigateToFile ".\snippet.html"
    
    driver.ImplicitWait = 1000
    
    'split the table cells into two groups Red and Blue
    With driver.FindElements(By.TagName, "td")
        For i = 1 To .Count
            If i Mod 2 = 0 Then
                'ordinal position in collection is even
                elemsBlue.Add .Item(i)
            Else
                'ordinal position in collection is odd
                elemsRed.Add .Item(i)
            End If
        Next i
    End With
    
    driver.Wait 1500
    
    'highlight the Blue group
    elemsBlue.Highlight borderSizePx:=2, borderColor:=Blue, ScrollIntoView:=False
    
    driver.Wait 1500
    
    'highlight the Red group
    elemsRed.Highlight borderSizePx:=2, borderColor:=Red, ScrollIntoView:=False, unHighlightLast:=False
    
    driver.Wait 1500
    
    'unhighlight the Blue group
    elemsBlue.UnHighlight
    
    driver.Wait 1500
    
    'unhighlight the Red group
    elemsRed.UnHighlight
    
    driver.Wait 1500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
