;Note: using v1 as Teams is currently incompatible with v2
#include <Teams>

GroupAdd, firefox, ahk_class MozillaWindowClass ; Add only Internet Explorer windows to this group.
return ; End of autoexecute section.

ProgramFilesX86 = C:\Program Files (x86)
SetTitleMatchMode, 2

EnvGet, UserProfile, UserProfile
DocsDir = %UserProfile%\STFC\Documents
MusicDir = %UserProfile%\Music

PythonEnv = %UserProfile%\Miniconda3\envs\py313
AnacondaCommand = %UserProfile%\Miniconda3\python.exe %UserProfile%\Miniconda3\cwp.py %PythonEnv% %PythonEnv%\python.exe

;Esc to quit Calc, Snipping Tool, and Notepad
#IfWinActive ahk_class CalcFrame
Esc::!F4
#IfWinActive ahk_exe notepad.exe
Esc::!F4
#IfWinActive ahk_class Microsoft-Windows-Tablet-SnipperEditor
Esc::!F4

;#IfWinActive ahk_class MozillaWindowClass
;\::Send {Alt Down}{Left}{Alt Up}

;activate Excel/Firefox/Outlook/Word/PPT if they exist, otherwise start them
#IfWinExist ahk_class XLMAIN
#x::WinActivate

#IfWinExist ahk_class MozillaWindowClass
;!\::
;    WinActivate
;    Send {Alt Down}d{Alt Up}`%
;    Return

#a::GroupActivate, firefox, r
#IfWinExist ahk_exe OUTLOOK.EXE
#o::WinActivate
#IfWinExist ahk_class OpusApp
#w::WinActivate
#IfWinExist ahk_class PPTFrameClass
#q::WinActivate
#IfWinExist Python
#c::WinActivate 

;shortcuts for timesheet entry page
#IfWinActive,Time Entry:
Down::Send {TAB 15}
Up::Send {Shift Down}{TAB 15}{Shift Up}

;Alt-D to highlight 'address bar' (name box) in Excel
#IfWinActive ahk_class XLMAIN
!d::ControlFocus Edit1
;shift+wheel for horizontal scrolling in Excel
+WheelDown::ComObjActive("Excel.Application").ActiveWindow.SmallScroll(0,0,2,0)
+WheelUp::ComObjActive("Excel.Application").ActiveWindow.SmallScroll(0,0,0,2)

#IfWinNotActive ahk_class XLMAIN  ;still want to use F4 in Excel
F4:: ;toggle video for Teams and Zoom
    Teams_Video()
    ControlSend, , {F4}, Zoom Meeting
    Return

#IfWinActive

F1:: ;mute/unmute both Teams and Zoom
    Teams_Mute()
    ControlSend, , {F1}, Zoom Meeting
    Return
    
#z::Send {Media_Play_Pause}
#+z::Send {Media_Next}
!F1::Send {Volume_Mute}

;launch the Zoom meeting that's nearest to now in the calendar
#^z::Run %AnacondaCommand% %DocsDir%\Scripts\join_zoom_meeting.pyw
;write notes for the Zoom meeting that's nearest to now in the calendar
#^n::Run %AnacondaCommand% %DocsDir%\Scripts\start_meeting_notes.pyw

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
;Ctrl-Win-T to toggle window transparency https://superuser.com/questions/272812/simple-transparency-toggle-with-autohotkey#answer-272955
#^t::
    WinGet, currentTransparency, Transparent, A
    if (currentTransparency = OFF)
    {
        WinSet, Transparent, 150, A
    }
    else
    {
        WinSet, Transparent, OFF, A
    }
    return
;Win-` to move window to other monitor
;^`::Send #+{Left}{Ctrl Up}
#b::Run %UserProfile%
#1::Run C:\
#2::Run %DocsDir%
#+w::Run mailto:
#i::Run %ProgramFiles%\irfanview\i_view64.exe
#n::Run C:\ProgramData\chocolatey\lib\metapad\tools\metapad.exe
#c::Run cmd.exe "/K" title Python & cd %UserProfile%\Misc\Scripts & %UserProfile%\Miniconda3\Scripts\activate.bat %PythonEnv% & python
#t::Run taskmgr.exe
; Win-Shift-T for a random Trello task
#+t::Run %AnacondaCommand% %DocsDir%\Scripts\oddjob.py
#+r::Run %AnacondaCommand% %UserProfile%\Misc\scripts\random_cd.py
#o::Run OUTLOOK.EXE
#w::Run WINWORD.EXE /q
#q::Run POWERPNT.EXE /s
#x::Run EXCEL.EXE /e
#+g::Run %MusicDir%
; jump through some hoops to get Remote Desktop app working (scaling works better with this than mstsc)
; see https://answers.microsoft.com/en-us/windows/forum/windows_10-windows_store/starting-windows-10-store-app-from-the-command/836354c5-b5af-4d6c-b414-80e40ed14675
#+x::Run explorer.exe shell:appsFolder\Microsoft.RemoteDesktop_8wekyb3d8bbwe!App
#`::WinMinimize,a

; to be used on Remote Desktop when Win key isn't available - snap windows to top/bottom/left/right using Alt+Numpad2/4/6/8
; !Numpad4::Send #{Left}
; !Numpad8::Send #{Up}
; !Numpad6::Send #{Right}
; !Numpad2::Send #{Down}

; names with accents
::Dusan::Dušan
::Topalovic::Topalović
::Pockar::Počkar

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
::CO2e::CO₂e

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

+#d::toggleDarkMode()

toggleDarkMode()
{
    static key := "", mode
    if !key
        RegRead mode, % key := "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", SystemUsesLightTheme
    mode ^= 1
    RegWrite REG_DWORD, % key, AppsUseLightTheme   , % mode
    RegWrite REG_DWORD, % key, SystemUsesLightTheme, % mode
}

