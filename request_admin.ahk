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

; generate the custom event that runs Chocolatey update
; see https://qtechbabble.wordpress.com/2021/09/09/use-system-events-to-trigger-administrator-scheduled-tasks-from-a-standard-user-account/
; source: BenShepherdCustomEvent
; id: 447870928072
Run, powershell -command "Write-EventLog -LogName Application -Source 'BenShepherdCustomEvent' -EntryType Information -EventId 8072 -Message 'Admin session started.'", , Hide
