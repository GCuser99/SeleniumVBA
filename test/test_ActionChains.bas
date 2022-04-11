Attribute VB_Name = "test_ActionChains"
Sub test_action_chain()
    Dim driver As New WebDriver, actions As ActionChain
    
    Dim from1 As WebElement, to1 As WebElement
    Dim from2 As WebElement, to2 As WebElement
    Dim from3 As WebElement, to3 As WebElement
    Dim from4 As WebElement, to4 As WebElement
    Dim elem As WebElement
    
    driver.StartChrome
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    Set from1 = driver.FindElement(by.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(by.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(by.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(by.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(by.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(by.XPath, "//*[@id='amt8']/li")
    
    driver.Wait 1000
    
    Set actions = driver.ActionChain
    actions.ScrollBy , 500
    actions.DragAndDrop(from1, to1).Wait
    actions.DragAndDrop(from2, to2).Wait
    actions.DragAndDrop(from3, to3).Wait
    'an alternative method to Drag and Drop
    actions.ClickAndHold(from4).MoveToElement(to4).ReleaseButton.Wait (1000)
    actions.Perform 'do all the actions defined above
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

Sub test_drag_and_drop()
    Dim driver As New WebDriver
    Dim from1 As WebElement, to1 As WebElement
    Dim from2 As WebElement, to2 As WebElement
    Dim from3 As WebElement, to3 As WebElement
    Dim from4 As WebElement, to4 As WebElement
    
    driver.StartChrome
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    Set from1 = driver.FindElement(by.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(by.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(by.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(by.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(by.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(by.XPath, "//*[@id='amt8']/li")
    
    driver.ScrollTo , 500
    
    'WebDriver and WebElement DragAndDrop made from action chains
    from1.DragAndDrop to1
    driver.Wait 500
    from2.DragAndDrop to2
    driver.Wait 500
    from3.DragAndDrop to3
    driver.Wait 500
    from4.DragAndDrop to4
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
    
End Sub

