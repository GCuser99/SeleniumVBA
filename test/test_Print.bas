Attribute VB_Name = "test_Print"
Sub test_print()
    Dim driver As New WebDriver
    Dim settings As New PrintSettings
    
    driver.StartChrome
    
    'must open browser in headless (invisible) mode for PrintToPDF to work
    driver.OpenBrowser , True
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.name, "q").SendKeys "This is COOL!" & vbCrLf
    
    driver.Wait 1000
    
    settings.Units = svbaInches
    settings.MarginsAll = 0.4
    settings.Orientation = svbaPortrait
    settings.PrintScale = 1
    'settings.PageRanges "1-2"  'prints the first 2 pages
    'settings.PageRanges 1, 2   'prints the first 2 pages
    settings.PageRanges 2       'prints only the 2nd page
    
    driver.PrintToPDF , settings

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

Sub test_screenshot()
    Dim driver As New WebDriver
    Dim keys As New Keyboard
    Dim caps As Capabilities
    Dim params As New Dictionary
    
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
    Dim keys As New Keyboard
    Dim caps As Capabilities
    Dim params As New Dictionary
    
    driver.StartChrome

    driver.OpenBrowser
    
    driver.NavigateTo "https://www.google.com/"
    driver.Wait 1000
    
    driver.FindElement(by.name, "q").SendKeys "This is COOL!" & vbCrLf
    
    driver.FindElement(by.name, "q").SaveScreenshot

    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub
