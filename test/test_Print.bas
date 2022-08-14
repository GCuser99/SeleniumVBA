Attribute VB_Name = "test_Print"
Sub test_print()
    Dim driver As New WebDriver
    Dim settings As New WebPrintSettings
    
    driver.DefaultIOFolder = ThisWorkbook.Path

    driver.StartChrome
    'must open browser in headless (invisible) mode for PrintToPDF to work
    driver.OpenBrowser , True
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.Name, "q").SendKeys "This is COOL!" & vbCrLf
    
    driver.Wait 1000
    
    settings.Units = svbaInches
    settings.MarginsAll = 0.4
    settings.Orientation = svbaPortrait
    settings.PrintScale = 1
    'settings.PageRanges "1-2"  'prints the first 2 pages
    'settings.PageRanges 1, 2   'prints the first 2 pages
    'settings.PageRanges 2       'prints only the 2nd page
    
    'prints pdf file to specified filePath parameter (defaults to .\printpage.pdf)
    driver.PrintToPDF , settings

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_screenshot()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    Dim caps As WebCapabilities
    Dim params As New Dictionary
    
    driver.DefaultIOFolder = ThisWorkbook.Path
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.SaveScreenshot

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_screenshot()
    Dim driver As New WebDriver
    Dim keys As New WebKeyboard
    Dim caps As WebCapabilities
    Dim params As New Dictionary
    
    driver.DefaultIOFolder = ThisWorkbook.Path
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.Name, "q").SendKeys "This is COOL!" & vbCrLf
    driver.Wait 1000
    driver.FindElement(by.Name, "q").SaveScreenshot

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
