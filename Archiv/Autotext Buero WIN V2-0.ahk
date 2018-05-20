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
; :C1:HOTSTRING::  Grosskleinschreibung des Hotstrings bleibt immer gleich
; :*:HOTSTRING:: - führt den String ohne abschließendes Leerzeichen sofort aus
; :?;HOTSTRING:: - erkennt den String auch innerhalb eines Wortes: blablaHOTSTRINGblabla
; :O:HOTSTRING:: requiring an ending character, but delete the ending char after insert "AK " -> "ARNE", not "ARNE "


; Weitere Modifiert für Hotstrings sind hier erklärt https://autohotkey.com/docs/Hotstrings.htm

; Start Init Section 

SetTitleMatchMode, 2
SetKeyDelay, 20

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

return ; End Init Section

:*:[ANM::[ANMERKUNG] 
:*:@aw::arne.wingbermuehle@opuscapita.com
:*:@arne::arnewingbermuehle@gmail.com
; :*:arne.::arne.wingbermuehle@opuscapita.com
:*:arne.wingbermuehle@o::arne.wingbermuehle@opuscapita.com
:C1*:#aw::Arne Wingbermühle
:C1*:#arne::Arne Wingbermühle
:C1:aw::Arne Wingbermühle
:C1*:Arne W::Arne Wingbermühle




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


:C1*:hz,::
(
Hallo zusammen,


)

:OC1*:hh,::
send Hallo Herr NAME,
send {LEFT 5} ; Cursor for NAME 
send +{RIGHT 4} ; selektiere NAME
return

:C1:#toc:: TODO OpusCapita: 
:C1:#toi:: TODO ifm: 
:*:todo o::TODO OpusCapita: 
:*:todo i::TODO ifm: 



::#protokoll::
(
Hallo zusammen,

anbei das bereits abgestimmte Protokoll für den Jour Fix vom 31.12.2019.
)

:O:mfg::
(
Gruss
Arne Wingbermühle
)

:O:lga::
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

::#AK::Akzeptanzkritieren:
::#to::Tagesorganisation

;SHIFT ALT CTRL H
+!^h::run "C:\Program Files (x86)\Notepad++\notepad++.exe" "%A_ScriptDir%\%A_ScriptName%"

;SHIFT ALT CTRL R
!^r::run "C:\Program Files (x86)\phyxcalc\phyxcalc.exe"


#IfWinActive, Notepad
~^s::
	Sleep 250
	Reload
	MsgBox, You pressed Ctrl S to Save in Notepad and Reload the script
Return
#IfWinActive

; PRJ TicketIDs im Browser in Jira öffnen
; benötigt das konfigurierte Searchkeyword im Firefox "jira"
#IfWinActive, ahk_class MozillaWindowClass
^!j::
	Send  ^c
	sleep 250
	Send  ^l  ; activate location bar in firefox
	sleep 250
	Send jira{Space}
	sleep 250	
	send ^v{ENTER}
	; MsgBox, Firefox Ctrl J fertig
Return
#IfWinActive


; PRJ TicketIDs in Excel in Jira öffnen
; benötigt das konfigurierte Searchkeyword im Firefox "jira"
; XLMAIN ist Excel
#IfWinActive, ahk_class XLMAIN
^!j::
	Send  ^c
	sleep 250
	IfWinNotExist ahk_class MozillaWindowClass
	{
		MsgBox, Firefox läuft nicht -> ich starte Firefox
		run, C:\Program Files (x86)\Mozilla Firefox\firefox.exe
		return
	}
	; WinWaitActive, ahk_class MozillaWindowClass, , 5
	WinActivate, ahk_class MozillaWindowClass
	sleep 250
	if ErrorLevel
	{
		MsgBox, Firefox Fenster ist nicht aktivierbar, leider.
		return
	}
	Send  ^l  ; activate location bar in firefox
	sleep 250
	Send jira{Space}
	sleep 250	
	send ^v{ENTER}
	; MsgBox, Firefox Ctrl J fertig
Return
#IfWinActive


; Clipboard ohne jede Formatierung einfügen, Leerzeilen entfernen
^!v:: ; CTRL+ALT+V
	;msgbox "nonformattingpaste"
	ClipSaved := 	 ;save original clipboard contents
	clipboard = %clipboard% ;remove formatting
	clipboard := RegexReplace( Clipboard, "^\s+|\s+$" )
	Send ^v ;send the Ctrl+V command
	clipboard := ClipSaved ;restore the original clipboard contents
	ClipSaved := ;clear the variable
Return





