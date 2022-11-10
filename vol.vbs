MsgBox "Warning:Fatle ERROR", vbCritical,"Error"
WScript.Sleep 300
do While True = True
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys(chr(&hAD))
    
    WScript.Sleep 600
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys(chr(&hAF))
loop