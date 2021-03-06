﻿ProgramFilesX86 = C:\Program Files (x86)
SetTitleMatchMode, 2

DocsDir = %UserProfile%\Documents
MusicDir = %UserProfile%\Music

;Esc to quit Calc, Snipping Tool, and Notepad
#IfWinActive ahk_class CalcFrame
Esc::!F4
#IfWinActive ahk_exe notepad.exe
Esc::!F4
#IfWinActive ahk_class Microsoft-Windows-Tablet-SnipperEditor
Esc::!F4

;activate Excel/Firefox/Outlook/Word/PPT if they exist, otherwise start them
#IfWinExist ahk_class XLMAIN
#x::WinActivate
#IfWinExist ahk_class MozillaWindowClass
#a::WinActivate
#IfWinExist ben.shepherd@stfc.ac.uk - Outlook
#o::WinActivate
#IfWinExist ahk_class OpusApp
#w::WinActivate
#IfWinExist ahk_class PPTFrameClass
#q::WinActivate
#IfWinExist Anaconda Prompt
#c::WinActivate 

; #IfWinExist, Google Play Music
; #z::ControlSend, ahk_parent, {SPACE}, ahk_exe firefox.exe

;shortcuts for timesheet entry page
#IfWinActive,Time Entry:
Down::Send {TAB 15}
Up::Send {Shift Down}{TAB 15}{Shift Up}

;Alt-D to highlight 'address bar' (name box) in Excel
#IfWinActive ahk_class XLMAIN
!d::ControlFocus Edit1

;Caps Lock to start/stop OBS recording if PPT slide show active
;#IfWinActive ahk_class screenClass
;CapsLock::ControlSend, , a, OBS

#IfWinActive

#z::Send {Media_Play_Pause}
#+z::Send {Media_Next}

;launch the Zoom meeting that's nearest to now in the calendar
#^z::Run C:\Users\bjs54\Miniconda3\pythonw.exe C:\Users\bjs54\Documents\Scripts\join_zoom_meeting.pyw
;if there are two at the same time, join the second one
#^+z::Run C:\Users\bjs54\Miniconda3\pythonw.exe C:\Users\bjs54\Documents\Scripts\join_zoom_meeting.pyw second
;write notes for the Zoom meeting that's nearest to now in the calendar
#^n::Run C:\Users\bjs54\Miniconda3\pythonw.exe C:\Users\bjs54\Documents\Scripts\join_zoom_meeting.pyw notes
;if there are two at the same time, make notes for the second one
#^+n::Run C:\Users\bjs54\Miniconda3\pythonw.exe C:\Users\bjs54\Documents\Scripts\join_zoom_meeting.pyw second notes

;Ctrl-Win-` to toggle window always on top
#^`::
	WinSet,AlwaysOnTop,Toggle,A
	WinGetTitle, Title, A
	WinGet, ExStyle, ExStyle, A
		if (ExStyle & 0x8)  ; 0x8 is WS_EX_TOPMOST.
			OnTop = on top
		else
			OnTop = not on top
	TrayTip, AutoHotKey, Set %Title% %OnTop%, 30
	Return

;Win-` to move window to other monitor
^`::Send #+{Left}{Ctrl Up}
#^\::Run c:\users\bjs54\Links\Recent Places.lnk
;#a::Run %ProgramFilesx86%\Mozilla Firefox\firefox.exe
#b::Run %UserProfile%
#1::Run C:\
#2::Run %DocsDir%
#+w::Run mailto:
#i::Run %ProgramFiles%\irfanview\i_view64.exe
;#v::Run %ProgramFiles%\irfanview\i_view64.exe ;Win-I is Settings on Win10
#n::Run C:\ProgramData\chocolatey\lib\metapad\tools\metapad.exe
#^p::Run cmd.exe
#c::Run cmd.exe "/K" C:\Users\bjs54\Miniconda3\Scripts\activate.bat C:\Users\bjs54\Miniconda3
#t::Run taskmgr.exe
#+r::Run C:\Users\bjs54\Miniconda3\python.exe %UserProfile%\Misc\scripts\random_cd.py
#o::Run OUTLOOK.EXE
#w::Run WINWORD.EXE /q
;#+a::Run C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe
#q::Run POWERPNT.EXE /s
#x::Run EXCEL.EXE /e
#g::Run %MusicDir%
; jump through some hoops to get Remote Desktop app working (scaling works better with this than mstsc)
; see https://answers.microsoft.com/en-us/windows/forum/windows_10-windows_store/starting-windows-10-store-app-from-the-command/836354c5-b5af-4d6c-b414-80e40ed14675
#+x::Run explorer.exe shell:appsFolder\Microsoft.RemoteDesktop_8wekyb3d8bbwe!App
#`::WinMinimize,a

; to be used on Remote Desktop when Win key isn't available - snap windows to top/bottom/left/right using Alt+Numpad2/4/6/8
; !Numpad4::Send #{Left}
; !Numpad8::Send #{Up}
; !Numpad6::Send #{Right}
; !Numpad2::Send #{Down}

;Greek letters
::\alpha::α
::\beta::β
::\gamma::γ
::\delta::δ
::\epsilon::ε
::\zeta::ζ
::\eta::η
::\theta::θ
::\iota::ι
::\kappa::κ
::\lambda::λ
::\mu::μ
::\nu::ν
::\xi::ξ
::\omicron::ο
::\pi::π
::\rho::ρ
::\sigma::σ
::\tau::τ
::\upsilon::υ
::\phi::φ
::\chi::χ
::\psi::ψ
::\omega::ω

;Subscripts
::CO2::CO₂
::kgCO2::kgCO₂
::kgCO2e::kgCO₂e
::tCO2e::tCO₂e

;Emojis
:::)::☺

;Ctrl+; to insert date, Ctrl+Shift+; to insert time (like Excel), Ctrl+Alt+; to insert yyyy-mm-dd
^;::
FormatTime, CurrentDateTime,, d/M
SendInput %CurrentDateTime%
return
^+;::
FormatTime, CurrentDateTime,, HH:mm
SendInput %CurrentDateTime%
return
^!;::
FormatTime, CurrentDateTime,, yyyy-MM-dd
SendInput %CurrentDateTime%
return

;Shift + Windows + Up (maximize a window across all displays) https://stackoverflow.com/a/9830200/470749
+#Up::
    WinGetActiveTitle, Title
    WinRestore, %Title%
   SysGet, X1, 76
   SysGet, Y1, 77
   SysGet, Width, 78
   SysGet, Height, 79
   WinMove, %Title%,, X1, Y1, Width, Height
return
