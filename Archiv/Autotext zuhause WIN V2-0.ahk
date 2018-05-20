#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force
; #Warn  ; Enable warnings to assist with detecting common errors.

; ! 	{Alt} 	{Alt down} 	{Alt up} 	Send !a presses ALT+a
; + 	{Shift} 	{Shift down} 	{Shift up} 	Send +abC sends the text "AbC"
; Send !+a presses ALT+SHIFT+a
; ^ 	{Ctrl} 	{Ctrl down} 	{Ctrl up} 	Send ^{Home} presses CONTROL+HOME
; # 	{LWin}
; {RWin} 	{LWin down}
; {RWin down} 	{LWin up}
; {RWin up} 	Send #e holds down the Windows key and then presses the letter "e"
; ~ Send the hotkey also to the app, instead of consuming it in autohokey

SetTitleMatchMode, 2
SetKeyDelay, 20

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

:*:@aw::arne.wingbermuehle@opuscapita.com
:*:@arne::arnewingbermuehle@gmail.com
:*:@coa::www.coming-of-agile.com
; :*:arne.::arne.wingbermuehle@opuscapita.com
:*:arne.wingbermuehle@o::arne.wingbermuehle@opuscapita.com
:C1*:#aw::Arne Wingbermühle
:C1*:#arne::Arne Wingbermühle
:C1:aw::Arne Wingbermühle
:C1:Arne W::Arne Wingbermühle


::pimbug::
(
h2. Scenario
I enter a term in Internet Explorer 11

h2. Result
If I save a term, PIM added whitespaces

h2. Steps to reproduce
# Open terms
Edit a term
Enter a new language term
Press "Save"

h2. Expected result
Only entered term is displayed without any whitespaces
)


::mfg::
(
Gruss
Arne Wingbermühle
)

::lga::
(
Gruss
Årne
)

:C1:OC::OpusCapita

::#story::
(
As a USER I want FUNCTION so I can GOAL.

Aceptance Criteria:
- Nice UI

Documentation:
- short / long / not needed

Out of Scope:
- worlddomination
)


::#AK::Acceptance criteria:
::#to::Tagesorganisation


+!^h::run "D:\Program Files (x86)\Notepad++\notepad++.exe" "%A_ScriptDir%\%A_ScriptName%"
+!^k::run "C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe"


#IfWinActive, Notepad
~^s::
	Sleep 250
	Reload
	MsgBox, You pressed Ctrl S to Save in Notepad and Reload the script
Return
#IfWinActive

#IfWinActive, ahk_class MozillaWindowClass
^!j::
	Send  ^c
	clipboard = %clipboard% ;remove formatting
	sleep 250
	Send  ^l  ; activate location bar in firefox
	sleep 250
	Send jira{Space}
	sleep 250	
	send ^v{ENTER}
	; MsgBox, Firefox Ctrl J fertig
Return
#IfWinActive

^!v:: ; CTRL+ALT+V
;msgbox "nonformattingpaste"
ClipSaved := ClipboardAll ;save original clipboard contents
clipboard = %clipboard% ;remove formatting
Send ^v ;send the Ctrl+V command
Clipboard := ClipSaved ;restore the original clipboard contents
ClipSaved = ;clear the variable
Return

