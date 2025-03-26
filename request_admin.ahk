; start local admin session
Run "C:\Program Files (x86)\FastTrack Software\Admin By Request\AdminByRequest.exe" /Elevate
WinWait, ahk_exe AdminByRequest.exe
WinActivate, ahk_exe AdminByRequest.exe
Send, Chocolatey software update
Send, !o  ; OK button
Sleep, 1000
WinActivate, ahk_exe AdminByRequest.exe
Send, !o  ; OK button
