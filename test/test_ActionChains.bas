Attribute VB_Name = "test_ActionChains"
Sub test_action_chain()
    Dim Driver As New WebDriver, actions As ActionChain
    
    Dim from1 As WebElement, to1 As WebElement
    Dim from2 As WebElement, to2 As WebElement
    Dim from3 As WebElement, to3 As WebElement
    Dim from4 As WebElement, to4 As WebElement
    Dim elem As WebElement
    
    Driver.StartChrome
    
    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    Set from1 = Driver.FindElement(by.XPath, "//*[@id='credit2']/a")
    Set to1 = Driver.FindElement(by.XPath, "//*[@id='bank']/li")
    
    Set from2 = Driver.FindElement(by.XPath, "//*[@id='credit1']/a")
    Set to2 = Driver.FindElement(by.XPath, "//*[@id='loan']/li")
    
    Set from3 = Driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to3 = Driver.FindElement(by.XPath, "//*[@id='amt7']/li")
    
    Set from4 = Driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to4 = Driver.FindElement(by.XPath, "//*[@id='amt8']/li")
    
    Driver.Wait 1000
    
    Set actions = Driver.ActionChain
    actions.ScrollBy , 500
    actions.DragAndDrop(from1, to1).Wait
    actions.DragAndDrop(from2, to2).Wait
    actions.DragAndDrop(from3, to3).Wait
    'an alternative method to Drag and Drop
    actions.ClickAndHold(from4).MoveToElement(to4).ReleaseButton.Wait (1000)
    actions.Perform 'do all the actions defined above
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

Sub test_drag_and_drop()
    Dim Driver As New WebDriver
    Dim from1 As WebElement, to1 As WebElement
    Dim from2 As WebElement, to2 As WebElement
    Dim from3 As WebElement, to3 As WebElement
    Dim from4 As WebElement, to4 As WebElement
    
    Driver.StartChrome
    
    Driver.OpenBrowser
    
    Driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    Set from1 = Driver.FindElement(by.XPath, "//*[@id='credit2']/a")
    Set to1 = Driver.FindElement(by.XPath, "//*[@id='bank']/li")
    
    Set from2 = Driver.FindElement(by.XPath, "//*[@id='credit1']/a")
    Set to2 = Driver.FindElement(by.XPath, "//*[@id='loan']/li")
    
    Set from3 = Driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to3 = Driver.FindElement(by.XPath, "//*[@id='amt7']/li")
    
    Set from4 = Driver.FindElement(by.XPath, "//*[@id='fourth']/a")
    Set to4 = Driver.FindElement(by.XPath, "//*[@id='amt8']/li")
    
    Driver.ScrollTo , 500
    
    'WebDriver and WebElement DragAndDrop made from action chains
    from1.DragAndDrop to1
    Driver.Wait 500
    from2.DragAndDrop to2
    Driver.Wait 500
    from3.DragAndDrop to3
    Driver.Wait 500
    from4.DragAndDrop to4
    Driver.Wait 2000
    
    Driver.CloseBrowser
    Driver.Shutdown
    
End Sub

