Attribute VB_Name = "test_Wait"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

Sub test_ImplicitMaxWait()
    Dim driver As SeleniumVBA.WebDriver
    
    Set driver = SeleniumVBA.New_WebDriver

    driver.StartChrome
    driver.OpenBrowser
    
    driver.ImplicitMaxWait = 10000
    
    driver.NavigateTo "https://www.selenium.dev/selenium/web/ajaxy_page.html"
    
    driver.FindElementByName("typer").SendKeys "Hello New World!"
    driver.FindElementByID("red").Click
    driver.FindElementByName("submit").Click
    
    'wait for element creation...
    Debug.Assert driver.FindElementByClassName("label").GetText = "Hello New World!"
    
    driver.Wait 1000
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilDisplayed()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with a script to hide the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<head><title>Test Wait Until Displayed</title></head>" & _
    "<body>" & _
    "<div id='testDiv'>I'm ready now after " & timeDelay & " ms!</div>" & _
    "<script>" & _
    "  var content = document.getElementById('testDiv');" & _
    "  content.style.display='none';" & _
    "  setTimeout(function(){" & _
    "    content.style.display='inline';" & _
    "  }, " & timeDelay & ");" & _
    "</script>" & _
    "</body>" & _
    "</html>"

    driver.NavigateToString html
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Assert driver.IsDisplayed(elem) = False
    
    'wait for it to display...
    driver.WaitUntilDisplayed elem
    
    Debug.Assert driver.IsDisplayed(elem) = True
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilNotDisplayed()
    Dim driver As SeleniumVBA.WebDriver
    Dim elem As SeleniumVBA.WebElement
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with a script to block the display of an element
    
    html = "<!DOCTYPE html>" & _
    "<html>" & _
    "<head><title>Test Wait Until Not Displayed</title></head>" & _
    "<body>" & _
    "<div id='testDiv'>I'm displayed for " & timeDelay & " ms...</div>" & _
    "<script>" & _
    "  var content = document.getElementById('testDiv');" & _
    "  content.style.display='inline';" & _
    "  setTimeout(function(){" & _
    "    content.style.display='none';" & _
    "  }, " & timeDelay & ");" & _
    "</script>" & _
    "</body>" & _
    "</html>"

    driver.NavigateToString html
    
    'find the "not displayed" element
    Set elem = driver.FindElement(By.ID, "testDiv")
    
    Debug.Assert driver.IsDisplayed(elem) = True
    
    'wait for it to disappear...
    driver.WaitUntilNotDisplayed elem
    
    Debug.Assert driver.IsDisplayed(elem) = False
    
    'WaitUntilNotDisplayed allows for method chaining too
    'Debug.Print "Is displayed?:", driver.WaitUntilNotDisplayed(elem).IsDisplayed
    
    driver.Wait 500
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitUntilNotPresent()
    Dim driver As SeleniumVBA.WebDriver
    Dim html As String
    Dim timeDelay As Long
    
    Set driver = SeleniumVBA.New_WebDriver
    
    driver.StartChrome
    driver.OpenBrowser
    
    timeDelay = 3000
    
    'create an html with an element of interest that is removed after a delay
    html = vbNullString
    html = html & "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "    <head>" & vbCrLf
    html = html & "        <title>Is Not Present?</title>" & vbCrLf
    html = html & "        <script>" & vbCrLf
    html = html & "            function removeElementWithDelay(delay, elementId) {" & vbCrLf
    html = html & "                setTimeout(function() {" & vbCrLf
    html = html & "                const elementToRemove = document.getElementById(elementId);" & vbCrLf
    html = html & "                if (elementToRemove) {" & vbCrLf
    html = html & "                    elementToRemove.remove();" & vbCrLf
    html = html & "                }" & vbCrLf
    html = html & "                }, delay);" & vbCrLf
    html = html & "            }" & vbCrLf
    html = html & "        </script>" & vbCrLf
    html = html & "    </head>" & vbCrLf
    html = html & "    <body onLoad=""removeElementWithDelay(" & timeDelay & ", 'div here');"">" & vbCrLf
    html = html & "        <div id=""div here"">I'm here...</div>" & vbCrLf
    html = html & "    </body>" & vbCrLf
    html = html & "</html>"

    driver.NavigateToString html
    
    'wait until the div is removed
    driver.WaitUntilNotPresent By.ID, "div here"

    driver.Wait 500
        
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitForDownload()
    Dim driver As SeleniumVBA.WebDriver
    Dim caps As SeleniumVBA.WebCapabilities
    
    Set driver = SeleniumVBA.New_WebDriver
   
    driver.StartChrome
    
    'set the directory path for saving download to
    Set caps = driver.CreateCapabilities
    caps.SetDownloadPrefs ".\"
    driver.OpenBrowser caps
    
    'delete legacy copy if it exists
    driver.DeleteFiles ".\test.pdf"
    
    driver.NavigateTo "https://github.com/GCuser99/SeleniumVBA/raw/main/dev/test_files/test.pdf"
    
    'wait until the download is complete
    driver.WaitForDownload ".\test.pdf"
    
    driver.DeleteFiles ".\test.pdf"
    
    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitForJQuery()
    Dim driver As New WebDriver
    Dim t As Single, errNum As Long
    Dim html As String

    html = "<!DOCTYPE html><html><head><title>jQuery Fake Transport Test</title>"
    html = html & "<script src=""https://code.jquery.com/jquery-3.7.1.min.js""></script>"
    html = html & "</head><body>"
    html = html & "<button id=""myButton"">Start AJAX</button>"
    html = html & "<p id=""output"">Ready</p>"
    html = html & "<script>"

    ' Fake transport: any $.ajax call with a fakeDelay option completes after
    ' that many ms with no network traffic. fakeStatus sets the HTTP status (default 200).
    ' Calls without fakeDelay fall through to the real XHR transport.
    html = html & "$.ajaxTransport('+text', function (o) {"
    html = html & "  if (o.fakeDelay == null) return;"
    html = html & "  var t;"
    html = html & "  return {"
    html = html & "    send: function (h, complete) {"
    html = html & "      t = setTimeout(function () {"
    html = html & "        complete(o.fakeStatus || 200, 'OK', { text: 'ok' });"
    html = html & "      }, o.fakeDelay);"
    html = html & "    },"
    html = html & "    abort: function () { clearTimeout(t); }"
    html = html & "  };"
    html = html & "});"

    ' Helper so test steps stay short: fake(ms) or fake(ms, {extra options})
    html = html & "window.fake = function (ms, extra) {"
    html = html & "  return $.ajax($.extend({ url: '/fake', dataType: 'text', fakeDelay: ms }, extra));"
    html = html & "};"

    ' Show state from jQuery's global events instead of a polling interval,
    ' so #output is correct when the waiter returns
    html = html & "$(document).on('ajaxStart', function () { $('#output').text('busy'); });"
    html = html & "$(document).on('ajaxStop', function () { $('#output').text('jQuery.active = ' + jQuery.active); });"

    html = html & "$(function () {"
    html = html & "  $('#myButton').click(function () { fake(1000); });"
    html = html & "});"

    html = html & "</script></body></html>"

    driver.StartChrome
    driver.OpenBrowser
    
    ' 1. Button click: single 1000 ms request; #output updated by ajaxStop
    driver.NavigateToString html
    driver.FindElementByID("myButton").Click
    driver.WaitForIdleJQuery
    Debug.Assert driver.FindElementByID("output").GetText = "jQuery.active = 0"

    ' 2. Three requests at once: all three must have completed
    driver.NavigateToString html
    driver.ExecuteScript "window.n = 0; [100, 400, 200].forEach(function (d) { fake(d, { complete: function () { window.n++; } }); });"
    driver.WaitForIdleJQuery settleTime:=100
    Debug.Assert driver.ExecuteScript("return n") = 3

    ' 3. Timeout: must raise the error, and quickly (not after the full 3 s)
    driver.NavigateToString html
    driver.ExecuteScript "fake(3000);"
    t = Timer
    On Error Resume Next
    driver.WaitForIdleJQuery 500
    errNum = Err.Number
    On Error GoTo 0
    Debug.Assert errNum = 404
    Debug.Assert Timer - t < 1.5

    ' 4. Chained request after a 50 ms gap: settle window must catch the second one
    driver.NavigateToString html
    driver.ExecuteScript "window.secondDone = false; fake(200, { success: function () { setTimeout(function () { fake(200, { success: function () { window.secondDone = true; } }); }, 50); } });"
    driver.WaitForIdleJQuery settleTime:=100
    Debug.Assert driver.ExecuteScript("return secondDone") = True

    ' 5. HTTP 500: counter still drops, request reports an error status
    driver.NavigateToString html
    driver.ExecuteScript "window.st = ''; fake(200, { fakeStatus: 500, complete: function (x, s) { window.st = s; } });"
    driver.WaitForIdleJQuery
    Debug.Assert driver.ExecuteScript("return st") = "error"
    Debug.Assert driver.ExecuteScript("return jQuery.active") = 0

    ' 6. Aborted request: counter drops when the request is aborted
    driver.NavigateToString html
    driver.ExecuteScript "window.st = ''; var x = fake(1000, { complete: function (x, s) { window.st = s; } }); setTimeout(function () { x.abort(); }, 100);"
    driver.WaitForIdleJQuery
    Debug.Assert driver.ExecuteScript("return st") = "abort"

    ' 7. global:false is NOT tracked: waiter returns before the request finishes (documents the limitation)
    driver.NavigateToString html
    driver.ExecuteScript "window.gDone = false; fake(300, { global: false, success: function () { window.gDone = true; } });"
    driver.WaitForIdleJQuery
    Debug.Assert driver.ExecuteScript("return gDone") = False

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitForIdleDOM()
    Dim driver As New WebDriver
    Dim html As String, errNum As Long, t As Double

    html = "<!DOCTYPE html><html><head><title>DOM Idle Test</title>"
    html = html & "<script src=""https://code.jquery.com/jquery-3.7.1.min.js""></script>"
    html = html & "<style>#panel{display:none;height:200px;background:#eee}</style>"
    html = html & "</head><body>"
    html = html & "<p>Clock: <span id=""clock"">0</span></p>"
    html = html & "<button id=""startClock"">Start clock</button>"
    html = html & "<button id=""showPanel"">Show panel</button>"
    html = html & "<button id=""buildList"">Build list</button>"
    html = html & "<div id=""content""><div id=""panel"">Panel</div><ul id=""list""></ul></div>"
    html = html & "<script>"
    ' Clock outside #content: updates every 50 ms, forever
    html = html & "$('#startClock').click(function () { var n = 0; setInterval(function () { $('#clock').text(++n); }, 50); });"
    ' jQuery animation: changes #panel's style attribute every frame for 1 s
    html = html & "$('#showPanel').click(function () { $('#panel').slideDown(1000); });"
    ' Chunked render: one <li> every 20 ms, 30 in total, no network
    html = html & "$('#buildList').click(function () { var n = 0; var h = setInterval(function () {"
    html = html & "  $('#list').append('<li>Item ' + (++n) + '</li>'); if (n === 30) clearInterval(h); }, 20); });"
    html = html & "</script></body></html>"

    driver.StartChrome
    driver.OpenBrowser

    ' 1. jQuery animation: wait until slideDown has finished
    driver.NavigateToString html
    driver.FindElementByID("showPanel").Click
    Debug.Assert driver.ExecuteScript("return $('#panel').is(':animated')") = True   ' still animating
    driver.WaitForIdleDOM 500, 5000
    Debug.Assert driver.ExecuteScript("return $('#panel').is(':animated')") = False
    Debug.Assert driver.ExecuteScript("return $('#panel').height()") = 200

    ' 2. Chunked render: wait until all 30 items are in
    driver.NavigateToString html
    driver.FindElementByID("buildList").Click
    Debug.Assert driver.ExecuteScript("return $('#list li').length") < 30           ' still rendering
    driver.WaitForIdleDOM 100, 5000
    Debug.Assert driver.ExecuteScript("return $('#list li').length") = 30

    ' 3a. Ticking clock + chunked render, wait limited to #content: clock is ignored
    driver.NavigateToString html
    driver.FindElementByID("startClock").Click
    driver.FindElementByID("buildList").Click
    driver.FindElementByID("content").WaitForIdleDOM 100, 5000
    Debug.Assert driver.ExecuteScript("return $('#list li').length") = 30

    ' 3b. Same page, wait on the whole document: the clock never stops, so it must time out
    t = Timer
    On Error Resume Next
    driver.WaitForIdleDOM 100, 1000
    errNum = Err.Number
    On Error GoTo 0
    Debug.Assert errNum = 404
    Debug.Assert Timer - t < 3                                                        ' gave up near 1 s, not 30 s
    
    html = "<!DOCTYPE html><html><head><title>Shadow DOM Idle Test</title></head><body>"
    html = html & "<button id=""build"">Build list</button>"
    html = html & "<div id=""host""></div>"
    html = html & "<script>"
    ' Open shadow root on #host containing an empty <ul id='list'>
    html = html & "var sr = document.getElementById('host').attachShadow({ mode: 'open' });"
    html = html & "var ul = document.createElement('ul'); ul.id = 'list'; sr.appendChild(ul);"
    ' Button adds 30 <li> inside the shadow root, one every 30 ms (about 0.9 s in total)
    html = html & "document.getElementById('build').addEventListener('click', function () {"
    html = html & "  var n = 0;"
    html = html & "  var h = setInterval(function () {"
    html = html & "    var li = document.createElement('li'); li.textContent = 'Item ' + (++n); ul.appendChild(li);"
    html = html & "    if (n === 30) clearInterval(h);"
    html = html & "  }, 30);"
    html = html & "});"
    html = html & "</script></body></html>"

    '----------------------- ShadowRoot test ----------------------
    ' 1. Watching the whole document does NOT see changes inside the shadow root:
    '    the waiter returns early, before the list is complete (documents the limitation)
    Dim sr As WebShadowRoot
    driver.NavigateToString html
    driver.FindElementByID("build").Click
    driver.WaitForIdleDOM 100, 5000
    Debug.Assert driver.ExecuteScript("return document.getElementById('host').shadowRoot.querySelectorAll('li').length") < 30

    ' 2. Watching the shadow root itself sees the changes: waits for all 30 items
    driver.NavigateToString html
    Set sr = driver.FindElementByID("host").GetShadowRoot
    driver.FindElementByID("build").Click
    driver.WaitForIdleDOM 100, 5000, sr
    Debug.Assert driver.ExecuteScript("return document.getElementById('host').shadowRoot.querySelectorAll('li').length") = 30

    ' 3. Timeout inside the shadow root: rendering takes about 0.9 s, so a 300 ms limit must raise 404
    driver.NavigateToString html
    Set sr = driver.FindElementByID("host").GetShadowRoot
    driver.FindElementByID("build").Click
    On Error Resume Next
    driver.WaitForIdleDOM 100, 300, sr
    errNum = Err.Number
    On Error GoTo 0
    Debug.Assert errNum = 404

    driver.CloseBrowser
    driver.Shutdown
End Sub

Sub test_WaitForIdleNetworkAndDOM()
    Dim driver As New WebDriver
    Dim html As String, r As String, t As Double
    Dim errNum As Long
    html = "<!DOCTYPE html><html><head><title>Idle State Test</title></head><body>"
    html = html & "<p id=""status"">idle</p>"
    html = html & "<script>"
    ' Fake network: delays come from the URL. ms = time until response, body = extra time to read
    ' the body, hang = never responds. The hooks wrap these fakes exactly as they would the real ones.
    html = html & "function q(u, k) { var m = new RegExp(k + '=(\\d+)').exec(u); return m ? +m[1] : 0; }"
    html = html & "function setStatus(t) { document.getElementById('status').textContent = t; }"
    html = html & "window.Response = function (bodyMs) { this.ok = true; this.bodyMs = bodyMs; };"
    html = html & "Response.prototype.json = function () { var ms = this.bodyMs; return new Promise(function (r) { setTimeout(function () { r({}); }, ms); }); };"
    html = html & "window.fetch = function (u) { return new Promise(function (r) { if (u.indexOf('hang') >= 0) return; setTimeout(function () { r(new Response(q(u, 'body'))); }, q(u, 'ms')); }); };"
    html = html & "XMLHttpRequest.prototype.open = function (m, u) { this._u = u; };"
    html = html & "XMLHttpRequest.prototype.send = function () { var x = this; setTimeout(function () { x.dispatchEvent(new Event('load')); x.dispatchEvent(new Event('loadend')); }, q(x._u, 'ms')); };"
    html = html & "function xhr(ms, onload) { var x = new XMLHttpRequest(); x.open('GET', '/x?ms=' + ms); x.onload = onload; x.send(); }"
    html = html & "</script></body></html>"
    driver.StartChrome
    driver.OpenBrowser

    ' A. fetch resolves when headers arrive (300 ms); body read takes another 300 ms, then render
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "fetch('/a?ms=300&body=300').then(function (r) { return r.json(); }).then(function () { setStatus('A done'); });"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "A done"

    ' B. Long-poll that never returns is ignored via the ignore list; the normal request is still waited for
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks (Array("hang"))
    driver.ExecuteScript "fetch('/poll?hang'); fetch('/b?ms=200').then(function (r) { return r.json(); }).then(function () { setStatus('B done'); });"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "B done"

    ' C. XHR, render in onload
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "xhr(300, function () { setStatus('C done'); });"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "C done"

    ' D. Updates that only change a text node's content (20 updates, 30 ms apart)
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "var t = document.getElementById('status').firstChild, n = 0; var h = setInterval(function () { t.data = String(++n); if (n === 20) clearInterval(h); }, 30);"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "20"

    ' E. Chained requests with a 50 ms gap and no DOM change in between
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "xhr(300, function () { setTimeout(function () { xhr(300, function () { setStatus('E done'); }); }, 50); });"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "E done"

    ' F. Three requests at once
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "window.n = 0; [100, 400, 200].forEach(function (d) { fetch('/f?ms=' + d).then(function () { if (++n === 3) setStatus('F done'); }); });"
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert driver.FindElementByID("status").GetText = "F done"

    ' G. Request that never returns, with no ignore list: must time out, and promptly
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    driver.ExecuteScript "fetch('/poll?hang');"
    On Error Resume Next
    driver.WaitForIdleNetworkAndDOM 100, 1000
    Debug.Assert Err.Number = 404
    Debug.Assert Err.Description = "Error in WaitForIdleNetworkAndDOM: Maximum time exceeded while waiting for idle state"
    On Error GoTo 0

    ' H. Hooks not installed (as after a click that navigates to a new page)
    driver.NavigateToString html
    On Error Resume Next
    driver.WaitForIdleNetworkAndDOM 100, 1000
    Debug.Assert Err.Number = 405
    Debug.Assert Err.Description = "Error in WaitForIdleNetworkAndDOM: Hooks not installed on this page. Call InstallNetworkAndDOMHooks after navigating and before the action to wait on"
    On Error GoTo 0

    ' I. Nothing happening: returns after about the idle time
    driver.NavigateToString html
    driver.InstallNetworkAndDOMHooks
    t = Timer
    driver.WaitForIdleNetworkAndDOM 100, 5000
    Debug.Assert Timer - t < 0.15
    driver.CloseBrowser
    driver.Shutdown
End Sub
