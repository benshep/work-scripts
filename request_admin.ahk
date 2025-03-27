EnvGet, UserProfile, UserProfile
DocsDir = %UserProfile%\STFC\Documents

; select a reason for requesting admin access
filePath = %DocsDir%\Scripts\admin_reasons.txt
FileRead, fileContent, %filePath%
sentences := StrSplit(fileContent, "`n")
Random, randomIndex, 1, % sentences.MaxIndex()
randomSentence := sentences[randomIndex]

; start local admin session
Run "C:\Program Files (x86)\FastTrack Software\Admin By Request\AdminByRequest.exe" /Elevate
WinWait, ahk_exe AdminByRequest.exe
WinActivate, ahk_exe AdminByRequest.exe
Send, %randomSentence%
Send, !o  ; OK button
Sleep, 1000
WinActivate, ahk_exe AdminByRequest.exe
Send, !o  ; OK button
