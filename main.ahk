; Boilerplate
#NoEnv ; performance and compatibility
#Warn  ; catching common errors.
SetWorkingDir %A_ScriptDir%  ; consistent starting directory.

; Configuration saving
INI_FILE := "ahk.ini"  ; Persistent script variables
NONE_VALUE := "NONE"   ; No script variable can be this value

; Libraries
#Include lib\saleslogix.ahk

; (Optional) Additional shortcuts which would not work for everyone
#Include *i usa.ahk
#Include *i p.nanda.ahk

;----------------------------------------------------------------------
; [Windows Key + w] List RMA report in web browser
;----------------------------------------------------------------------
#w::
IniRead, REGION, %INI_FILE%, RMA, region, NONE_VALUE
RMAShowReport() {
  RMA_BROWSER_TITLE = RMA Report - Google Chrome
  Run http://intranet/scripts/wip/rma.asp
  WinWait %RMA_BROWSER_TITLE%,,20
  if ErrorLevel
  return
  Global REGION
  WinActivate
  Send {Tab}{Tab}{Enter}  ; Click "Regional Search"
  if REGION = AndorUK
  {
    Send {Tab}{Tab}{Down}  ; Select region "AndorUK"
  }
  else if REGION = AndorUS
  {
    Send {Tab}{Tab}{Down}{Down}  ; Select region "AndorUS"
  }
  else if REGION = AndorJapan
  {
    Send {Tab}{Tab}{Down}{Down}{Down}  ; Select region "AndorJapan"
  }
  else
  {
    Send {Tab}{Tab}{Down}{Down}{Down}{Down}  ; Select region "All"
  }
  Send +{Tab}{Down}{Down}  ; Select status "All"
  Send {Tab}{Tab}{Enter}   ; Click "Go"
  return
}
if REGION = NONE_VALUE
{
  Gui, Add, Text,, Select your region:
  Gui, Add, DropDownList, vREGION, AndorUS||AndorUK|AndorJapan|All
  Gui, Add, Button, Default, OK
  Gui, Show
  return
  
  GuiClose:
  ButtonOK:
  Gui, Submit
  ; MsgBox You entered %REGION%
  IniWrite %REGION%, %INI_FILE%, RMA, region
  if ErrorLevel
  {
    MsgBox Error: Could not write %REGION% to %INI_FILE%
    exit
  }
  RMAShowReport()
  return
}
; Region was already present in the INI_FILE
RMAShowReport()
return

;----------------------------------------------------------------------
; [Windows Key + m] Launch WIP RMA database
;----------------------------------------------------------------------
#m::
Send ^c
; FIXME: Login credentials are assumed to be  correct.  If login details
;        are wrong the user would need to edit the INI file.
IniRead, RMA_USER, %INI_FILE%, RMA, username, NONE_VALUE
if RMA_USER = NONE_VALUE
{
  InputBox RMA_USER, New RMA user, Enter RMA username
  IniWrite %RMA_USER%, %INI_FILE%, RMA, username
  if ErrorLevel
  {
    MsgBox Error: Could not write %RMA_USER% to %INI_FILE%
    return
  }
}
IniRead, RMA_PASS, %INI_FILE%, RMA, password, NONE_VALUE
if RMA_PASS = NONE_VALUE
{
  InputBox RMA_PASS, New RMA user, Enter RMA password, HIDE
  IniWrite %RMA_PASS%, %INI_FILE%, RMA, password
  if ErrorLevel
  {
    MsgBox Error: Could not write your password to %INI_FILE%
    return
  }
}
Run C:\vb6\Andor\Andor.exe
WinWait, Login,,20
if ErrorLevel
  return
WinActivate
Send %RMA_USER%{tab}%RMA_PASS%{tab}{Enter}

WinWait, Andor Technology,,10
if ErrorLevel
  return

WinActivate
Send {Down}{Enter}{Down}{Down}{Enter}

WinWait, Andor (Live),,10
if ErrorLevel
  return

WinActivate, Andor (Live)
begin := RegExMatch(clipboard, "R\d\d\d\d\d")
if begin != 0
{
  rma := SubStr(clipboard, begin, 6)
  Send %rma%{tab}
  ; Prevent clipboard data from clobbering subsequent RMA searches
  clipboard = ; clear clipboard
}
return

;----------------------------------------------------------------------
; [Windows Key + d] Date paste
;----------------------------------------------------------------------
#d::
IniRead, INITIALS, %INI_FILE%, Timestamp, initials, NONE_VALUE
if INITIALS = NONE_VALUE
{
  InputBox INITIALS, New Timestamp user, Enter your initials
  IniWrite %INITIALS%, %INI_FILE%, Timestamp, initials
  if ErrorLevel
  {
    MsgBox Error: Could not write %INITIALS% to %INI_FILE%
    return
  }
}
TimeVar := A_Now
FormatTime, TimeVar, A_Now, %INITIALS% ddd MMM dd, yyyy
Send %TimeVar%
return

;----------------------------------------------------------------------
; [Windows Key + o] Sales order search from clipboard
;----------------------------------------------------------------------
#o::
Send ^c
Run http://intranet/cm.mccann/Sales Orders/

WinWait, Sales Orders,,10
if ErrorLevel
  return
WinActivate
Send {tab}^v{tab}{Enter}
clipboard = ; clear clipboard
return

;----------------------------------------------------------------------
; [Windows Key + b] BOM search from clipboard
;----------------------------------------------------------------------
#b::
Send ^c
Run http://intranet/bomreport/

WinWait, Shamrock Components,,20
if ErrorLevel
  return
WinActivate
while (A_Cursor = "AppStarting")
  continue
Send {tab}%clipboard%{tab}{Enter}
clipboard = ; clear clipboard
return

;----------------------------------------------------------------------
; [Windows Key + t] Ticket search from clipboard or Outlook e-mail title
;----------------------------------------------------------------------
#t::
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
    MsgBox,, Ticket search, No ticket number found in clipboard or e-mail title
    return
}
open_ticket(ticket)
return

;----------------------------------------------------------------------
; [Windows Key + a] Save attachments from Outlook to ticket folder
;----------------------------------------------------------------------
#a::
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
    MsgBox,, Import ticket attachments, No ticket number found in clipboard or e-mail title
    return
}
FileCreateDir, %A_Desktop%\%ticket%
; attachment selection window
Send !fna
WinWait,Save All Attachments,,0.1,
only_one_attachment := 0
if ErrorLevel
{
    ; only one attachment
    only_one_attachment := 1
    Send !fn{enter}
    WinWait,Save Attachment,,0.1,
    if ErrorLevel
        return
}
else
{
    WinActivate
    Send {enter}
    ; file save dialog
    WinWait,Save All Attachments,,0.1,
    if ErrorLevel
        return
}
Send {home}
clipboard = %A_Desktop%\%ticket%\
Send ^v{enter}
; exit hotscript if you get a file overwrite warning message
WinWait,Microsoft Office Outlook,,0.1,
if !ErrorLevel
{
    return
}
; TODO: add saved attachments to SalesLogix:
;       open_ticket(ticket)
;       loop add_attachment(file)
return