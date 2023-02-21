Set WshShell = CreateObject("WScript.Shell") 
 
iActive = WinWaitActivate "Title", 5 
If iActive = 1 Then WshShell.SendKeys("%{F4}") 
 
Function WinWaitActivate(Title, TimeOut) 
    TimerInit = Timer 
    iRet = WshShell.AppActivate(Title) 
    While iRet = 0 
        Wscript.Sleep(10) 
        iRet = WshShell.AppActivate(Title) 
        If TimeOut > 0 And (Timer - TimerInit) >= TimeOut Then 
            WinWaitActivate = 0 
            Exit Function 
        End If 
    WEnd 
    WinWaitActivate = 1 
End Function