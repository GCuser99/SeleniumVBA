'this kills any stranded WebDriver processes and associated child processes
'modified from SeleniumBasic (https://github.com/florentbr/SeleniumBasic/blob/master/Scripts/RunCleaner.vbs)

i = 0
driverList = Array("msedgedriver.exe", "chromedriver.exe", "geckodriver.exe", "IEdriverServer.exe")
queryConditional="Name='" & Join(driverList, "' Or Name='") & "'"
Set mgt = GetObject("winmgmts:")
On Error Resume Next
For Each p In mgt.ExecQuery("Select * from Win32_Process Where " & queryConditional)
    For Each cp In mgt.ExecQuery("Select * from Win32_Process Where ParentProcessId=" & p.ProcessId)
        cp.Terminate
    Next
    i = i + 1
    p.Terminate
Next

MsgBox "Done - " & i & " drivers killed"

