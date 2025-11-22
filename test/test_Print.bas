Attribute VB_Name = "test_Print"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_print()
    Dim driver As SeleniumVBA.WebDriver
    Dim settings As SeleniumVBA.WebPrintSettings
    Dim keys As SeleniumVBA.WebKeyboard

    Set driver = SeleniumVBA.New_WebDriver
    Set settings = SeleniumVBA.New_WebPrintSettings
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.FindElement(By.ID, "searchInput").SendKeys "Leonardo da Vinci" & keys.EnterKey
    
    'insure that the entire page is loaded using simulated scrolling
    'not needed in this particular case but could be useful in "lazy load"
    'situations such as in https://www.yahoo.co.jp/
    Dim scrollHeight As Long, lastScrollHeight As Long
    scrollHeight = driver.GetScrollHeight
    Do
        driver.ScrollToBottom enSpeed:=jump_instant
        driver.Wait 500
        lastScrollHeight = scrollHeight
        scrollHeight = driver.GetScrollHeight
    Loop Until scrollHeight = lastScrollHeight
    
    settings.Units = svbaInches
    settings.MarginsAll = 0.4
    settings.Orientation = svbaPortrait
    settings.PrintScale = 1
    
    settings.DisplayHeaderFooter = True
    'the following is close to the default header and footer - not needed but provided for example:
    'settings.HeaderTemplate = "<div style='font-size:10px; width:100%; display:flex; justify-content:space-between; align-items:center; padding:0 40px;'><span class='date' style='flex:1; text-align:left;'></span><span class='title' style='flex:1; text-align:center;'></span><span style='flex:1;'></span></div>"
    'settings.FooterTemplate = "<div style='font-size:10px; width:100%; display:flex; justify-content:space-between; padding:0 40px;'><span class='url'></span><span><span class='pageNumber'></span>/<span class='totalPages'></span></span></div>"
    
    'settings.PageRanges "1-2"  'prints the first 2 pages
    'settings.PageRanges 1, 2   'prints the first 2 pages
    'settings.PageRanges 2       'prints only the 2nd page
    
    'prints pdf file to specified filePath parameter (defaults to .\printpage.pdf)
    driver.PrintToPDF , settings

    driver.Wait 1000
    
    driver.DeleteFiles "printpage.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_screenshot()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.SaveScreenshot

    driver.Wait 1000
    
    driver.DeleteFiles "screenshot.png"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_screenshot_full()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.SaveScreenshot fullScreenShot:=True

    driver.Wait 1000
    
    driver.DeleteFiles "screenshot.png"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_element_screenshot()
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    'driver.DefaultIOFolder = ThisWorkbook.path '(this is the default)
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 1000
    
    driver.FindElement(By.ID, "searchInput").SendKeys "Leonardo da Vinci"
    driver.Wait 1000
    driver.FindElement(By.ID, "searchInput").SaveScreenshot

    driver.Wait 1000
    
    driver.DeleteFiles "screenshot.png"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
