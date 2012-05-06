; Shortcuts implemented in this script
;
; [Windows Key + w] List RMA report in web browser
; [Windows Key + m] Launch WIP RMA database
; [Windows Key + d] Date paste
; [Windows Key + o] Sales order search from clipboard
; [Windows Key + b] BOM search from clipboard
; [Windows Key + t] Ticket search from Clipboard or Outlook e-mail title
;

; Boilerplate
#NoEnv ; performance and compatibility
#Warn  ; catching common errors.
SetWorkingDir %A_ScriptDir%  ; consistent starting directory.

; Configuration saving
INI_FILE := "ahk.ini"  ; Persistent script variables
NONE_VALUE := "NONE"   ; No script variable can be this value

; (Optional) Additional shortcuts which would not work for everyone
#Include *i personal.ahk
#Include *i experimental.ahk

;----------------------------------------------------------------------
; [Windows Key + w] RMA report in web browser for US RMAs
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

; TODO: Backport Chome web page has mouse cursor load from personal.ahk instead 
;       of using a timeout
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
; [Windows Key + t] Ticket search from Clipboard or Outlook e-mail title
;----------------------------------------------------------------------
#t::
SetTitleMatchMode, Slow
SetTitleMatchMode, 2
match_found := 0
begin := RegExMatch(clipboard, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
if begin != 0
{
  ticket := SubStr(clipboard, begin, 13)
  ;MsgBox %ticket%
  match_found := 1
  ; Prevent this clipboard data from clobbering subsequent Outlook subject
  ; selections
  clipboard = ; clear clipboard
}
; Check highlighted Outlook e-mail title for ticket number
if match_found <> 1
{
  WinGetText, text, Microsoft Outlook
  ;MsgBox DEBUG: %text%
  Loop, Parse, text, `n, `r ; parses variable text by newline
  {
    if match_found <> 1
    {
      subject:=A_LoopField
      ;MsgBox DEBUG: %subject%
      begin := RegExMatch(subject, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
      if begin != 0
      {
        ticket := SubStr(subject, begin, 13)
        ;MsgBox %ticket%
        match_found := 1
      }
    }
  }
  VarSetCapacity(text,0) ; frees data from variable text
}
if match_found <> 1
  return
WinWait,SalesLogix,,10,Personal
if ErrorLevel
  return
WinActivate
; Gets rid of the Modal Window: "Sync Client logix has successfully applied
; transactions to your local database.  ...refresh the client now?"
WinWait, Confirm,,0.1
; Ignore ErrorLevel from WinWait ... No prompt Window, so nothing to do
if !ErrorLevel
{
  WinActivate
  Send !y
  ; Kludge: wait for SLX refresh to complete before finding ticket.
  ;         Otherwise sending keystrokes while the main toolbar reloads
  ;         creates an access violation and SLX has to be restarted.
  SetTitleMatchMode, 1
  SetTitleMatchMode, Fast
  counter = 10
  WinWait,Sage SalesLogix -,,5
  while ( ErrorLevel &&  counter > 0 )
  {
    WinWait,Sage SalesLogix -,,2
    counter--
  }
  Sleep,4000
}
Send !ltt%ticket%{tab}{enter}
WinWait, Lookup Ticket,,10
if ErrorLevel
  return
WinActivate
Send !o
return