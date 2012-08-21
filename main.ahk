; Boilerplate
#NoEnv ; performance and compatibility
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
wip_window_exists()
{
  WinWait, Andor (Live),,2
  if ErrorLevel
  {
    ; if the RMA view is closed, also close the RMA menu window
    ; to prevent double login
    WinClose, Andor Technology
    return 0
  }
  return 1
}
open_rma_in_existing_wip_window(close_existing_rma = 0)
{
  WinActivate, Andor (Live)
  if close_existing_rma
  {
    Send {Escape}
    WinWait, Andor (Live),,1
    if ErrorLevel
    {
      WinWait, Andor Technology,, 1
      WinActivate
      Send {Enter}
    }
    WinWait, Andor (Live),,1
    WinActivate
  }
  
  begin := RegExMatch(clipboard, "R\d\d\d\d\d")
  if begin != 0
  {
    rma := SubStr(clipboard, begin, 6)
    Send %rma%{tab}
    ; Prevent clipboard data from clobbering subsequent RMA searches
    clipboard = ; clear clipboard
  }
  return begin
}
if wip_window_exists()
{
  open_rma_in_existing_wip_window(1)
  return
}

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

if ErrorLevel
  return

WinActivate
Send {Down}{Enter}{Down}{Down}{Enter}

if !wip_window_exists()
  return

open_rma_in_existing_wip_window()
return

;----------------------------------------------------------------------
; [Windows Key + 5] Date paste
;----------------------------------------------------------------------
#5::
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
FormatTime, TimeVar, A_Now, ddd MMM dd, yyyy
Send %INITIALS% %TimeVar%
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
Send ^c
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
    MsgBox,, Ticket search, No ticket number found in clipboard or e-mail title
    return
}
open_ticket(ticket)
return

;----------------------------------------------------------------------
; [Windows Key + a] Save attachments from Outlook to Desktop folder
;----------------------------------------------------------------------
#a::
IniRead, OUTLOOK_ATTACH, %INI_FILE%, Outlook, attachments, NONE_VALUE
if OUTLOOK_ATTACH = NONE_VALUE
{
  FileSelectFolder, OUTLOOK_ATTACH, %A_Desktop%,
                  , Select folder to save Outlook attachments
  IniWrite %OUTLOOK_ATTACH%, %INI_FILE%, Outlook, attachments
  if ErrorLevel
  {
    MsgBox Error: Could not write %OUTLOOK_ATTACH% to %INI_FILE%
    return
  }
}
contact_name := get_contact_name_from_outlook_subject()
if contact_name = %NONE_VALUE%
{
    MsgBox,, Save Outlook attachments, No contact name found in e-mail title
    return
}
FileCreateDir, %OUTLOOK_ATTACH%\%contact_name%
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
clipboard = %OUTLOOK_ATTACH%\%contact_name%\
Send ^v{enter}
; exit hotscript if you get a file overwrite warning message
WinWait,Microsoft Office Outlook,,0.1,
if !ErrorLevel
{
    return
}
return

;----------------------------------------------------------------------
; [Windows Key + z] Bugzilla search
;----------------------------------------------------------------------
#z::
Send ^c
found := RegExMatch(clipboard, "(\d+)", bug)
if found
  Run http://uk00083/show_bug.cgi?id=%bug1%
return