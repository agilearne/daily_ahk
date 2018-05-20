; **************************
; Autohotkey Scriptsammlung Scrummaster PO Agile Coach
; gedacht für Autostart, tägliche Arbeit
; Author: Arne Wingbermühle
; **************************


; **************************
; Spezielle Script Direktiven für Autohotkey
; **************************
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance Force			; Script wird nur einmal ausgeführt quasi als Singleton
;#Warn  ; Enable warnings to assist with detecting common errors.



; **************************
; Doku: Encoding und Zeichensatz (Umlaute ...)
; ************************** 
; Encoding of Script: ANSI CP1251
; Für andere Codepages (GitHub Default! Unixe!) siehe auch: https://autohotkey.com/docs/Scripts.htm#cp



; **************************
; Doku: Präfix für Hotkeys
; **************************
;
; ! 	{Alt} 	{Alt down} 	{Alt up} 	Send !a presses ALT+a
; + 	{Shift} 	{Shift down} 	{Shift up} 	Send +abC sends the text "AbC"
; ^ 	{Ctrl} 	{Ctrl down} 	{Ctrl up} 	Send ^{Home} presses CONTROL+HOME
; Send ^!+a presses CTRL+ALT+SHIFT+a



; **************************
; Doku: Modifikatoren für Hotkeys
; **************************
; ~ Send the hotkey also to the app, instead of consuming it in autohokey
; Weitere Tastenkürzel und Modifikatoren https://autohotkey.com/docs/Hotkeys.htm
; Tastenliste: https://autohotkey.com/docs/KeyList.htm



; **************************
; Doku: Modifikatoren für die Hotstrings
; **************************

; :C1:HOTSTRING::  Grosskleinschreibung des Hotstrings bleibt immer gleich
; :*:HOTSTRING:: - führt den String ohne abschließendes Leerzeichen sofort aus
; :?;HOTSTRING:: - erkennt den String auch innerhalb eines Wortes: blablaHOTSTRINGblabla
; :O:HOTSTRING:: requiring an ending character, but delete the ending char after insert "AK " -> "ARNE", not "ARNE "
; Weitere Modifizierer für Hotstrings: https://autohotkey.com/docs/Hotstrings.htm



; **************************
; Doku: Debugging
; **************************
; Anstelle von Log oder Msgbox benutzt man hier eine richtigen Debugger, mit Breakpoints etc.
; Doku: https://autohotkey.com/docs/Scripts.htm#idebug
; Debugger: https://code.google.com/archive/p/xdebugclient/
; Aufruf der Scripte im Debug Modus:
; To enable interactive debugging, first launch a supported debugger client then launch the script with the /Debug command-line switch.
; (der Debugger kann auch nach dem Scriptstart gestartet werden, siehe Doku)
; AutoHotkey.exe /Debug=SERVER:PORT ...
; AutoHotkey /Debug "Hotstrings.ahk"



; **************************
; Start Init Section 
; **************************

SetTitleMatchMode, 2		; 2: A window's title can contain WinTitle anywhere inside it to be a match. 
SetKeyDelay, 20				; Sets the delay in milliseconds that will occur after each keystroke sent by Send and ControlSend

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Um Doubleclicks abzufangen (nur Doku, unbenutzt)
SysGet, SM_CXDOUBLECLK, 36
SysGet, SM_CYDOUBLECLK, 37
DoubleClickTime := DllCall("User32\GetDoubleClickTime")

; Testtickets für Jira JSON Integration
JiraURLBase := "http://support.jcatalog.com/rest/api/latest/"
JiraIssueID := "PRJ-33645"

; Library for JSON handling
;#include json\JSON.ahk

; Default für Change Tags in geänderten Dokumenten
ChangedVersion := ""


return 			; Sicheres Return, für Fehler bei den Hotstring etc.
; End auto-execute / init section




; **************************
; Hotstrings
; **************************


; **************************
; Indiviuelle Namen, Emails und Marken
; **************************

:*:@arne::arnewingbermuehle@gmail.com
:*:@aw::arnewingbermuehle@gmail.com
:*:@coa::arne@coming-of-agile.com
:*:arnewingbermuehle@g::arne.wingbermuehle@gmail.com
:*:arne@c::arne.wingbermuehle@coming-of-agile.com
:C1*:#coa::coming-of-agile.com
:C1*:#aw::Arne Wingbermühle
:C1*:#arne::Arne Wingbermühle
:C1:aw::Arne Wingbermühle
:C1*:Arne W::Arne Wingbermühle
:C1*:sonneng::sonnengleiter@gmx.de
:C1*:coa::coming-of-agile.com





; **************************
; Emails Templates
; **************************


; Grussformel Formell
:O:mfg::
(
Gruss
Arne Wingbermühle
)

; Grussformel Zwanglos
:O:lga::
(
Gruss
Årne
)


; Anrede Gruppe
:C1*:hz,::
(
Hallo zusammen,


)


; Anrede Einzelner Formel
:OC1*:hh,::
	send Hallo Herr NAME,
	send {LEFT 5} ; Cursor vor NAME 
	send +{RIGHT 4} ; selektiere NAME, so dass der überschrieben werden kann
return


; Versenden Protokoll Meeting Minutes 
::#protokoll::
(
Hallo zusammen,

anbei das Protokoll für den Meeting vom ?.
)



; **************************
; Templates für Tickets
; **************************

; Zitieren Block (Quote) für Jira
:O:#q::
:O:#quote::
	send {{}quote{}}{ENTER 2}{{}quote{}}{UP 1}
return

::#AK::Akzeptanzkritieren:

; Bugticket
::#bug::
(
h2. Scenario
I enter a term in Internet Explorer 11

h2. Result
If I save a document, the system does nothing

h2. Steps to reproduce
# Open document
Edit a document
Press "Save"

h2. Expected result
The documentchanges are saved
)


; Storyticket
::#story::
(
*As a* < type of user / PERSONA>
*I want* < some goal / action> 
*so that* < some reason / provide value>.

h2. Acceptance criteria
- Rule
- or Usecase
- or GIVEN WHEN THEN format

h2. Out of Scope
- make 10K in 10 sec with gold selling.

h2. How to demo
Click OK -> Worlddomination

h2. Documentation
- Script: worlddomination.groovy
- Go to app and login as admin/admin
- Creates a black hole. Sends in a spaceship with user and goes back in time. be the one and only emporer on earth. 
)



; **************************
; Templates für Protokolle und Zeiterfassung (überall)
; **************************

:*:[ANM::[ANMERKUNG]:
:C1:#tw::TODO Wingbermuehle: 
:*:Tageso::
:*:#tag::Tagesorganisation
return


; Setzen eines Update Textes, damit der in Zukunft automatisch ohne Tippen eingefügt werden kann (überall)
::#upd::
	if(ChangedVersion == "") {
		if setChangeVersionText("1.0", ChangedVersionEntry) {
			ChangedVersion := ChangedVersionEntry
		}
		else {
			return
		}
	}
	
	send [UPDATE V%ChangedVersion%]:
return

; Reset den VersionText so dass man ihn neu vergeben kann (überall)
; (und weil sich manche Leute keine Hotstrings merken können, sind es halt n anstelle von nur einem ...)
::#updset::
::#setupd::
::#updreset::
::#resetupd::
::#newupd::
::#updnew::
	ChangedVersionEntry := ""
	if setChangeVersionText(ChangedVersion, ChangedVersionEntry) {
		ChangedVersion := ChangedVersionEntry
	}
	else {
		return
	}
return


; setChangeVersionText: Eingabe oder Überschreiben einer Versionstextes
setChangeVersionText(DefaultText, ByRef VersionInput) {
	if(VersionInput == "") {
		InputBox VersionInput, Version eingeben, Version, , w, h, x, y, ,60,%DefaultText%
		if (ErrorLevel != 0) {
			VersionInput := ""
			return false
		}
	}
	return true
}




; **************************
; Templates für Protokolle (nur in GDocs)
; **************************

#IfWinActive, ahk_class MozillaWindowClass
#IfWinActive, Google Docs



; Fettmarkiertes TODO in GDocs für die verschiedenen Firmen etc.
:C1*:todo o::
	InsertBoldTodo("OTHER")
return

:C1*:todo i::
	InsertBoldTodo("INTERN")
return

:C1*:todo e::
	InsertBoldTodo("EXTERN")
return

InsertBoldTodo(Who) {
	send ^{b}TODO %who%:^{b}{SPACE}
}



; **************************
; Templates für Dokumentationen (nur in GDocs)
; **************************

; Update Präfix für Dokumentänderungen
:C1*O:#upd::
	if(ChangedVersion == "") {
		if setChangeVersionText("1.0", ChangedVersionEntry) {
			ChangedVersion := ChangedVersionEntry
		}
		else {
			return
		}
	}

	InsertChangePrefix(ChangedVersion)
return

; Update Präfix für Dokumentänderungen jeweils mit Version
:C1*O:#u11::
	InsertChangePrefix(1.1)
return

:C1*O:#u12::
	InsertChangePrefix(1.2)
return

:C1*O:#u13::
	InsertChangePrefix(1.3)
return

:C1*O:#u20::
	InsertChangePrefix(2.0)
return

:C1*O:#u21::
	InsertChangePrefix(2.1)
return

:C1*O:#u22::
	InsertChangePrefix(2.2)
return

:C1*O:#frage::
	send ^{b}[Frage an Kunde]^{b}:{SPACE}
return



; Function für das Einfügen von fettem Versionsänderungs-Tag in GDocs
InsertChangePrefix(Version) {
	send ^{b}[UPDATE V%Version%]^{b}:{SPACE}
}

#IfWinActive



; **************************
; Hotkeys
; **************************


; Alt Shift Ctrl: Diese Script Datei im Editor selber bearbeiten
+!^h::

	appExe := "\Notepad++\notepad++.exe"
	appPath := getAppPath(appExe)
	if (appPath) {
		run "%appPath%" "%A_ScriptDir%\%A_ScriptName%"
	}
	else {
		showSplash("Notepad++ nicht gefunden. Kann AHK Script nicht bearbeiten.", 3000)	
	}	
return



; Alt Ctrl N: Textdatei als Memo mit Timestampdateiname auf dem Desktop anlegen und bearbeiten
; Alt Ctrl M: Textdatei als Memo mit Timestampdateiname auf dem Desktop anlegen und bearbeiten
!^m::
!^n::
	const_Praefix_Memofile := "Memo "
	
	FormatTime, TimestampForFilename, , yyyyMMdd HH_mm		 
	FormatTime, TimestampForContent, , dd.MM.yyyy HH:mm

	Filename = %A_Desktop%\%const_Praefix_Memofile%%TimestampForFilename%.txt     	; A_Desktop: full path and name of the folder containing current user's desktop
	
	appExe := "\Notepad++\notepad++.exe"
	appPath := getAppPath(appExe)
	if (appPath) {
		FileAppend, %TimestampForContent%`n, %Filename%
		showSplash("Neue Textdatei unter " + Filename, 500)
		run "%appPath%" "%Filename%"
	}
	else {
		showSplash("Notepad++ nicht gefunden. Kann Memodatei nicht bearbeiten (und lege sie auch nicht an).", 3000)	
	}	
return



; Alt Ctrl R: Speziellen Taschenrechner phyxcalc öffnen
!^r::
	appExe := "\phyxcalc\phyxcalc.exe"
	appPath := getAppPath(appExe)
	if (appPath) {
		showSplash("Taschenrechner!", 500)
		SplitPath, appPath, ,workingDir
		run %appPath%, %workingDir%
		;run C:\Program Files (x86)\phyxcalc\phyxcalc.exe, C:\Program Files (x86)\phyxcalc\
	}
	else {
		showSplash("Physcalc ist nicht auffindbar.", 3000) 
	}	
return



; Alt Ctrl K: Keypass öffnen
!^k::
	appExe := "\KeePass Password Safe 2\KeePass.exe"
	appPath := getAppPath(appExe)
	if (appPath) {
		showSplash("KeePass!", 500)
		run %appPath%
	}
	else {
		showSplash("KeePass ist nicht auffindbar.", 3000) 
	}	
return



; Alt Ctrl T: Trello Website mit Personal Kanban öffnen
!^t::
	showSplash( "Öffne Trello im Browser ...", 400)
	Run, https://trello.com/b/dx0uLgio/personal-kanban
return



; PAUSE: Play / Pause Toggle beim Abspielen von Liedern - Mediaplayer (Spotify oder jeder andere Player, Windowsmediasteuerung)
; PAUSE: Starte Spotify wenn es noch nicht läuft
Pause::
{ 	
	Process, Exist, Spotify.exe ; check to see if Spotify.exe is running
	if !errorlevel {
		appExe := "\Spotify\Spotify.exe"
		appPath := getAppPath(appExe)
		if (appPath) {
			showSplash("Starte Spotify!", 500)
			run %appPath%
		}
		else {
			showSplash("Spotify ist nicht auffindbar.", 3000) 
			return
		}	
	}
	else {
		send {Media_Play_Pause}
	}
	return 
} 



; Ctrl PAUSE: Nächster Song - Mediaplayer (Spotify oder jeder andere Player, Windowsmediasteuerung)
; Doku autohotkey special mist: Break key. Since this is synonymous with Pause, use ^CtrlBreak in hotkeys instead of ^Pause or ^Break.
^CtrlBreak::
{
	send {Media_Next}
	return
}

; Strg Numpad+: Lautstärk nach oben (für 0815 Tastatur ohne Mediakeys)
^NumpadAdd::
{
	send {Volume_Up}
	return
}


; Strg Numpad-: Lautstärk nach unten (für 0815 Tastatur ohne Mediakeys)
^NumpadSub::
{
	send {Volume_Down}
	return
}
	


; Alt Ctrl V: Clipboard ohne jede Formatierung einfügen als Ascii
^!v::
	showSplash( "Clipboard als reines Ascii einfügen ...", 250)
		
	;ClipSaved := clipboard									;save original clipboard contents
	clipboard = %clipboard% 								;remove formatting
	; clipboard = RegexReplace( Clipboard, "^\s+|\s+$" )  	; Remove Linebreaks
	send ^v 												;send the Ctrl+V command
	;sleep 200
	;clipboard = %ClipSaved% 									;restore the original clipboard contents
	;ClipSaved = 											;clear the variable
return




; Ctrl S: Catch Save Command of this file in Notepad and reload script, so u save a few clicks... (nur in Notepad)
; only the scriptfile itself is reloaded not every file in notepad
; (catches hotkey but passes it through to Notepad)
#IfWinActive, Notepad
~^s::
	; Nur wenn Name des Scriptes im Notepad++ Windowtitel vorkommen, also das Script das aktive Fenster ist.
	if WinActive(A_ScriptName) {
		showSplash( "Reloading Script nach dem Speichern...", 400)
		reload
	}
		
Return
#IfWinActive




; **************************
; Hotkeys für Jira
; **************************

; Alt J: Abruf von Ticketdaten aus Jira (Test)
!j::
	;input_str := "key: value"
	;test_json := JSON.Dump(input_str)
	;msgbox %test_json%
	
	jira := new JIRAConnector(JiraBaseURL)
	jira := ""
return



; Ctrl Alt J; PRJ TicketIDs im Browser in Jira öffnen in einem neuen Tab (nur im Firefox)
; benötigt das konfigurierte Searchkeyword im Firefox "jira"
; Arbeitsweise: Ticket ID markieren -> Clipboardcopy -> neuer Tab mit der Suche in Jira via dem Keyword ...
#IfWinActive, ahk_class MozillaWindowClass
^!j::
	Send  ^c
	sleep 250
	send  ^t  ; open new tab (in foreground)
	send  ^l  ; activate location bar in firefox (locationbar text is selected)
	sleep 250
	Send jira{Space}
	sleep 250	
	send ^v{ENTER}
	; MsgBox, Firefox Ctrl J fertig
return
#IfWinActive



; PRJ TicketIDs in Excel in Jira öffnen (nur in Excel)
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
return
#IfWinActive




; Ctrl Alt Shift T: Wandelt eine Exceltabelle die bereits im Clipboard ist ins Jira MD Format und fügt diese an der aktuellen Stelle (sinnvollerweise also im Jira Ticket ...) ein
+!^t::
  StringReplace, ForFirstLine, Clipboard, `r`n, §
  StringSplit, ForFirstLine, ForFirstLine, §
  StringReplace, ForFirstLine1, ForFirstLine1, %A_Tab%, ||, All
  StringReplace, ForFirstLine2, ForFirstLine2, %A_Tab%, |, All
  ForFirstLine2 := RTrim(ForFirstLine2, "`r`n")
  StringReplace, ForFirstLine2, ForFirstLine2, `r`n, |`r`n|, All
  Clipboard := "||" . ForFirstLine1 . "||`r`n|" . ForFirstLine2 . "|"
  Send ^v
return



; **************************
; Utility Functions
; **************************



; getAppPath: finde eine Programmdatei an den typischen Orten unter Windows, weil es ja mehrere gibt, inkl. mehrerer Laufwerke
; gibt den gefundenen, qualifizierten Pfad inkl. Filename zurück. False wenn die App Datei nicht auffindbar ist.
; appExe: App Datei inkl. dem üblichen Installationsverzeichnis (Beispiel: "\Notepad++\notepad++.exe")
getAppPath(appExe) {

	; Laufwerksverzeichnis für das auffinden von Exe Dateien (nur erster Backslash, ohne letzten nach Verzeichniss)
	AppHarddrives := ["C:", "D:"]
	AppFolders := ["\Program Files", "\Program Files (x86)", "\Users\" A_UserName "\AppData\Local", "\Users\" A_UserName "\AppData\Roaming"]


	appExe := trim(appExe)
	if appExe == ""
		return False
		
	for indexHarddrives, Harddrive in AppHarddrives {
		for indexFolders, Folder in AppFolders {
			appFullPath := Harddrive Folder appExe
			
			; a := FileExist(appFullPath)
			IfExist, %appFullPath%
			{
				return %appFullPath%
			}		
		}	
	}
	return false
}



; showSplash: zeigt eine Nachricht als Splash für eine bestimmte Zeit in ms an (und wartet diese Zeit auch)
showSplash(displayText, splashWaitTime) {
	Progress, off
	Progress, B zh0 fs14 W600 CTBlack CWYellow, %displayText%
	Sleep %splashWaitTime%
	Progress, off
}

; Ermittelt die installierte Autohotkey Version 
getAHKVersion() {

	F1=http://ahkscript.org/download/1.1/version.txt
	;A_ver:=A_AHKVERSION

	;-if A_Is64bitOS

	info1 := "AHK Version:`t" A_AhkVersion
	   . "`nUnicode:`t`t" (A_IsUnicode ? "Yes " ((A_PtrSize=8) ? "(64-bit)" : "(32-bit)") : "No")
	   . "`nOperating System:`t" (!A_OSVersion ? A_OSType : A_OSVersion)
	   . "`nAdmin Rights:`t" (A_IsAdmin ? "Yes" : "No")

	if (A_IsUnicode)
	  codex := " Encoding is Unicode"
	else
	  codex := " Encoding is ANSI"
	info2:="Version: " ( InStr( (v:=A_AhkVersion), "1.1" ) ? "ahk_L " : "ahk_Basic " ) v  codex
	
	msgbox, ,AHK-Versions-Info, (Changed)`n%info1%`n------------------`n%info2%`n------------------
	return
}



; "Abstract" Class for "TicketSystems" like Jira, Rally, Trelle, LeanKit, ServiceNow, whatever
class TicketConnector {
	GetTicketData(TicketID) {
	}
}


/* Handles all the connection setting etc for a JIRA 6.4.11 onpremise installation
class JiraConnector extends TicketConnector{
	__New(BaseURL){
		return
    }

    __Delete(){
		return
    }
	
	GetTicketData(TicketID) {
	}
}