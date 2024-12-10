Attribute VB_Name = "test_ActionChains"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_action_chain()
    Dim driver As SeleniumVBA.WebDriver, actions As SeleniumVBA.WebActionChain
    Dim left1 As SeleniumVBA.WebElement, right1 As SeleniumVBA.WebElement
    Dim left2 As SeleniumVBA.WebElement, right2 As SeleniumVBA.WebElement
    Dim left3 As SeleniumVBA.WebElement, right3 As SeleniumVBA.WebElement
    Dim left4 As SeleniumVBA.WebElement, right4 As SeleniumVBA.WebElement
    Dim left5 As SeleniumVBA.WebElement, right5 As SeleniumVBA.WebElement
    Dim list1 As SeleniumVBA.WebElement, list2 As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartEdge
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/draggableLists.html"
    
    driver.Wait 500
    
    Set left1 = driver.FindElement(By.CssSelector, "#leftitem-1")
    Set left2 = driver.FindElement(By.CssSelector, "#leftitem-2")
    Set left3 = driver.FindElement(By.CssSelector, "#leftitem-3")
    Set left4 = driver.FindElement(By.CssSelector, "#leftitem-4")
    Set left5 = driver.FindElement(By.CssSelector, "#leftitem-5")
    
    Set right1 = driver.FindElement(By.CssSelector, "#rightitem-1")
    Set right2 = driver.FindElement(By.CssSelector, "#rightitem-2")
    Set right3 = driver.FindElement(By.CssSelector, "#rightitem-3")
    Set right4 = driver.FindElement(By.CssSelector, "#rightitem-4")
    Set right5 = driver.FindElement(By.CssSelector, "#rightitem-5")
    
    Set list1 = driver.FindElement(By.CssSelector, "#sortable1")
    Set list2 = driver.FindElement(By.CssSelector, "#sortable2")
    
    driver.Wait 500
    
    Set actions = driver.ActionChain
    actions.DragAndDrop left1, list2
    actions.DragAndDrop right1, list1
    actions.DragAndDrop left2, list2
    actions.DragAndDrop right2, list1
    actions.DragAndDrop left3, list2
    actions.DragAndDrop right3, list1
    actions.DragAndDrop left4, list2
    actions.DragAndDrop right4, list1
    actions.DragAndDrop left5, list2
    actions.DragAndDrop right5, list1
    actions.ClickAndHold(right3).MoveToElement(right5).ReleaseButton
    actions.ClickAndHold(right4).MoveToElement(right5).ReleaseButton
    actions.ClickAndHold(left4).MoveToElement(left5).ReleaseButton
    actions.Wait(1000).Perform 'do all the actions defined above
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_action_chain_sendkeys()
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim actions As SeleniumVBA.WebActionChain
    Dim textBox As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartEdge
    
    driver.OpenBrowser
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/single_text_input.html"
    driver.Wait 500
    
    Set textBox = driver.FindElement(By.ID, "textInput")
    
    Set actions = driver.ActionChain
    
    'build the chain and then execute with Perform method
    actions.MoveToElement(textBox).Click
    actions.KeyDown(keys.ShiftKey).SendKeys("upper case").KeyUp (keys.ShiftKey)
    actions.Perform

    driver.Wait 2000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub


