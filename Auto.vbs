Set Wshshell = CreateObject("Wscript.Shell")
WScript.Sleep 1000
WshShell.SendKeys"{ }"
Wscript.Sleep 4400
WshShell.SendKeys"{RIGHT}"
Wscript.Sleep 1000
WshShell.SendKeys"{LEFT}"REM'4.8s'
Wscript.Sleep 1000
WshShell.SendKeys"{DOWN}"REM'5.5s'
Wscript.Sleep 3200
Wshshell.SendKeys"{RIGHT}"REM'8.1s REAL 7.8'
Wscript.Sleep 2450
Wshshell.SendKeys"{RIGHT}"REM'10.4s'
Wscript.Sleep 300
Wshshell.SendKeys"{DOWN}"REM'10.6s'
Wscript.Sleep 190
Wshshell.SendKeys"{LEFT}"REM'11.0s'
Wscript.Sleep 200
Wshshell.SendKeys"{DOWN}"REM'11.2s-2s=9.2s'
Wscript.Sleep 500
Wshshell.SendKeys"{RIGHT}"REM'9.7s'
Wscript.Sleep 650
Wshshell.SendKeys"{DOWN}"REM'10.6s'
Wshshell.SendKeys"{F}"REM'10.6s'
Wscript.Sleep 1600
Wshshell.Sendkeys"{LEFT}"REM'11.6s'
Wscript.Sleep 400
Wshshell.Sendkeys"{DOWN}"REM'12.0s'
Wscript.Sleep 1500
Wshshell.SendKeys"{LEFT}"
Wscript.Sleep 350
Wshshell.SendKeys"{DOWN}"
Wscript.Sleep 2500
Wshshell.SendKeys"{RIGHT}"
Wscript.Sleep 500
Wshshell.SendKeys"{DOWN}"
Wscript.Sleep 2800
Wshshell.SendKeys"{LEFT}"
Wscript.Sleep 1200
Wshshell.SendKeys"{DOWN}"
msgbox"end"