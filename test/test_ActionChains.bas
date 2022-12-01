Attribute VB_Name = "test_ActionChains"
Option Explicit
Option Private Module

Sub test_action_chain()
    Dim driver As SeleniumVBA.WebDriver, actions As SeleniumVBA.WebActionChain
    Dim from1 As SeleniumVBA.WebElement, to1 As SeleniumVBA.WebElement
    Dim from2 As SeleniumVBA.WebElement, to2 As SeleniumVBA.WebElement
    Dim from3 As SeleniumVBA.WebElement, to3 As SeleniumVBA.WebElement
    Dim from4 As SeleniumVBA.WebElement, to4 As SeleniumVBA.WebElement
    Dim elem As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    driver.Wait 500
    
    Set from1 = driver.FindElement(By.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(By.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(By.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(By.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(By.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(By.XPath, "//*[@id='amt8']/li")
    
    driver.Wait 500
    
    Set actions = driver.ActionChain
    actions.ScrollBy , 500
    actions.DragAndDrop from1, to1
    actions.DragAndDrop from2, to2
    actions.DragAndDrop from3, to3
    actions.ClickAndHold(from4).MoveToElement(to4).ReleaseButton
    actions.Perform 'do all the actions defined above
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_action_chain_sendkeys()
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim actions As SeleniumVBA.WebActionChain
    Dim searchBox As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartEdge
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.wikipedia.org/"
    driver.Wait 500
    
    Set searchBox = driver.FindElement(By.ID, "searchInput")
    
    Set actions = driver.ActionChain
    
    'build the chain and then execute with Perform method
    actions.MoveToElement(searchBox).Click
    actions.KeyDown(keys.ShiftKey).SendKeys("upper case").KeyUp (keys.ShiftKey)
    actions.SendKeys(keys.ReturnKey).Perform

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_drag_and_drop()
    Dim driver As SeleniumVBA.WebDriver
    Dim from1 As SeleniumVBA.WebElement, to1 As SeleniumVBA.WebElement
    Dim from2 As SeleniumVBA.WebElement, to2 As SeleniumVBA.WebElement
    Dim from3 As SeleniumVBA.WebElement, to3 As SeleniumVBA.WebElement
    Dim from4 As SeleniumVBA.WebElement, to4 As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    driver.NavigateTo "https://demo.guru99.com/test/drag_drop.html"
    
    driver.Wait 500
    
    Set from1 = driver.FindElement(By.XPath, "//*[@id='credit2']/a")
    Set to1 = driver.FindElement(By.XPath, "//*[@id='bank']/li")
    
    Set from2 = driver.FindElement(By.XPath, "//*[@id='credit1']/a")
    Set to2 = driver.FindElement(By.XPath, "//*[@id='loan']/li")
    
    Set from3 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to3 = driver.FindElement(By.XPath, "//*[@id='amt7']/li")
    
    Set from4 = driver.FindElement(By.XPath, "//*[@id='fourth']/a")
    Set to4 = driver.FindElement(By.XPath, "//*[@id='amt8']/li")
    
    driver.ScrollTo , 500
    
    'WebDriver and WebElement DragAndDrop's method made from action chains
    from1.DragAndDrop to1
    from2.DragAndDrop to2
    from3.DragAndDrop to3
    from4.DragAndDrop to4
    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

