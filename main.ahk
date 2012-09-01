; Boilerplate
#NoEnv ; performance and compatibility
SetWorkingDir %A_ScriptDir%  ; consistent starting directory.

; Configuration saving
INI_FILE := "ahk.ini"  ; Persistent script variables
NONE_VALUE := "NONE"   ; No script variable can be this value

; Libraries
#Include lib\progress_bar.ahk
#Include lib\common.ahk
#Include lib\saleslogix.ahk
if A_OSVersion = WIN_XP
{
  #Include *i window_manager.ahk
}
; (Optional) Additional shortcuts which would not work for everyone
#Include *i usa.ahk
#Include *i p.nanda.ahk

;----------------------------------------------------------------------
; [Windows Key + w] List RMA report in web browser
;----------------------------------------------------------------------
#w::
create_progress_bar("RMA report")
add_progress_step("Reading locale")
add_progress_step("Opening default report")
add_progress_step("Getting local report")
step_progress_bar()
IniRead, REGION, %INI_FILE%, RMA, region, NONE_VALUE
RMAShowReport() {
  RMA_BROWSER_TITLE = RMA Report - Google Chrome
  step_progress_bar()
  Run http://intranet/scripts/wip/rma.asp
  WinWait %RMA_BROWSER_TITLE%,,60
  if ErrorLevel
    Goto, End_hotkey
  Global REGION
  WinActivate
  step_progress_bar()
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
  Return
}
if REGION = NONE_VALUE
{
  Gui, Add, Text,, Select your region:
  Gui, Add, DropDownList, vREGION, AndorUS||AndorUK|AndorJapan|All
  Gui, Add, Button, Default, OK
  Gui, Show
  Goto, end_hotkey
  
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
  Goto, end_hotkey
}
; Region was already present in the INI_FILE
RMAShowReport()
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + m] Launch WIP RMA database
;----------------------------------------------------------------------
#m::
create_progress_bar("Open RMA record")
copy_to_clipboard()
open_live_rma_window()
{
   WinActivate Andor Technology
   Send {Enter}
   WinWait, Andor (Live),,10
   WinActivate
}
wip_window_exists()
{
  IfWinExist, Andor (Live)
    Return 1
  Else
  {
    ; if the RMA view is closed, also close the RMA menu window
    ; to prevent double login
    WinClose, Andor Technology
    Return 0
  }
}
open_rma_in_existing_wip_window(close_existing_rma = 0)
{
  WinActivate, Andor (Live)
  if close_existing_rma
  {
    Send {Escape}
    ; if no RMA is active, this would close the Andor (Live) window,
    ; so we need to reopen it
    IfWinExist, Andor (Live)
      WinActivate
    Else
      open_live_rma_window()
  }

  begin := RegExMatch(clipboard, "R\d{5}")
  if begin != 0
    begin++  ; we don't want the R in the matched string
  else  ; fallback to checking just digits
    begin := RegExMatch(clipboard, "\d{5}")
  if begin != 0
  {
    rma := SubStr(clipboard, begin, 5)
    Send R%rma%{tab}
    ; Prevent clipboard data from clobbering subsequent RMA searches
  }
  Return 1
}
if wip_window_exists()
{
  if open_rma_in_existing_wip_window(1)
    Goto, end_hotkey
}
add_progress_step("Logging into database")
add_progress_step("Opening WIP window")
add_progress_step("Searching RMA record")
step_progress_bar()

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
    Goto, end_hotkey
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
    Goto, end_hotkey
  }
}
Run C:\vb6\Andor\Andor.exe
WinWait, Login,,20
if ErrorLevel
    Goto, end_hotkey
WinActivate
Send %RMA_USER%{tab}%RMA_PASS%{tab}{Enter}

if ErrorLevel
  Goto, end_hotkey

step_progress_bar()
WinActivate
Send {Down}{Enter}{Down}{Down}{Enter}

WinWait, Andor (Live),,10
if ErrorLevel
  Goto, end_hotkey

step_progress_bar()
open_rma_in_existing_wip_window()
Goto, end_hotkey

;----------------------------------------------------------------------
; [Ctrl + S] Save WIP RMA record
;----------------------------------------------------------------------
#IfWinActive Andor (Live)
^s::
  create_progress_bar("Save RMA record")
  add_progress_step("Extracting RMA number")
  add_progress_step("Saving RMA")
  add_progress_step("Reopening RMA")
  clipboard =
  ControlGetText, clipboard, ThunderRT6TextBox17 ; "RMA No" textbox
  if clipboard !=
  {
    step_progress_bar()
    ControlClick, ThunderRT6CommandButton6 ; "CR-Accept" button
    Send {Enter} ; since AHK click is too short for WIP GUI to register it

    ; Wait till the "RMA No" text field is in focus again
    Loop, 10
    {
      ControlGetFocus, control
      if control = ThunderRT6TextBox17
      {	
	step_progress_bar()
	Send ^v{Tab}
        Goto, End_hotkey
      }
      else
	Sleep, 500
    }
  }
  Goto, End_hotkey
#IfWinActive ; turn off context sensitivity

;----------------------------------------------------------------------
; [Ctrl + P] Open WIP RMA record in Microsoft Word for printing
;----------------------------------------------------------------------
#IfWinActive Andor (Live)
^p::
  create_progress_bar("Create RMA report")
  add_progress_step("Clicking 'Returns Print'")
  add_progress_step("Exporting to Word")
  add_progress_step("Closing 'Returns Print'")
  step_progress_bar()
  ControlClick, ThunderRT6CommandButton1 ; "Returns Print" button
  Send {Enter} ; since AHK click is too short for WIP GUI to register it
  
  get_rma_print_window()
  {
    ; Wait for untitled Print window
    SetTitleMatchMode, Slow
    WinWait,, 100`%, 10
    if ErrorLevel
      Return 0
    WinActivate
    Return 1
  }
  
  get_rma_export_window()
  {
    WinWait, Export,, 1
    If ErrorLevel
      Return 0
    WinActivate
    Return 1
  }
  
  if !get_rma_print_window()
    Goto, End_hotkey
  step_progress_bar()
  Loop, 10 ; loop since there's no way to know if the button has loaded
  {
    if !get_rma_export_window()
    {
      ControlClick, X320 Y40 ; Export button is at 326, 41
      Break
    }
  }
  
  Sleep, 500 ; sometimes the CSV window opens instead
  Send {Tab 2}{End}{Tab}{Home}{Enter}
  ; close export and print window
  step_progress_bar()
  if get_rma_export_window()
    Send {Esc}
  Sleep, 2000 ; otherwise Alt+F4 below doesn't work
  if get_rma_print_window()
    Send !{F4}
  Goto, End_hotkey
#IfWinActive ; turn off context sensitivity

;----------------------------------------------------------------------
; [Windows Key + 5] Date paste
;----------------------------------------------------------------------
#5::
create_progress_bar("Date stamp")
IniRead, INITIALS, %INI_FILE%, Timestamp, initials, NONE_VALUE
if INITIALS = NONE_VALUE
{
  InputBox INITIALS, New Timestamp user, Enter your initials
  IniWrite %INITIALS%, %INI_FILE%, Timestamp, initials
  if ErrorLevel
  {
    MsgBox Error: Could not write %INITIALS% to %INI_FILE%
    Goto, end_hotkey
  }
}
TimeVar := A_Now
FormatTime, TimeVar, A_Now, ddd MMM dd, yyyy
Send %INITIALS% %TimeVar%
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + o] Sales order search from clipboard
;----------------------------------------------------------------------
#o::
  create_progress_bar("Sales Order search")
  copy_to_clipboard()
  
  ; there's no native function to parse several regex matches, so one has to
  ; reuse the `begin` position parameter to check the full string
  begin = 1
  While begin := RegExMatch(clipboard, "([MX]?\d{6}[\/]?\d?)"
                            , match
                            , begin + StrLen(match))
  {
    sales_order%A_Index% := match1
    matches := A_Index
  }
  Loop, %matches%
  {
    add_progress_step("Opening web page")
    add_progress_step("Querying sales order")
  }
  
  ; open a new window for each match
  Loop, %matches%
  {
    step_progress_bar()
    Run http://intranet/cm.mccann/Sales Orders/
    step_progress_bar()
    Sleep, 500 ; wait for browser to clear the page title
    WinWait, Sales Orders,,10
    WinActivate
    clipboard := sales_order%A_Index%
    Send {Tab}^v{Tab}{Enter}
  }
    Goto, End_hotkey

;----------------------------------------------------------------------
; [Windows Key + b] BOM search from clipboard
;----------------------------------------------------------------------
#b::
create_progress_bar("BOM search")
add_progress_step("Opening web page")
add_progress_step("Querying part number")
copy_to_clipboard()
step_progress_bar()
Run http://intranet/bomreport/

WinWait, Shamrock Components,,40
if ErrorLevel
  Goto, end_hotkey
WinActivate
while (A_Cursor = "AppStarting")
  continue
step_progress_bar()
Send {tab}%clipboard%{tab}{Enter}
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + t] Ticket search from clipboard or Outlook e-mail title
;----------------------------------------------------------------------
#t::
create_progress_bar("Ticket search")
add_progress_step("Extracting Ticket ID")
add_progress_step("Opening in SalesLogix")
copy_to_clipboard()
step_progress_bar()
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
    MsgBox,, Ticket search, No ticket number found in clipboard or e-mail title
    Goto, end_hotkey
}
step_progress_bar()
open_ticket(ticket)
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + a] Save attachments from Outlook to Desktop folder
;----------------------------------------------------------------------
#a::
create_progress_bar("Save e-mail attachments")
add_progress_step("Creating folder")
add_progress_step("Getting attachments")
add_progress_step("Saving attachments")
IniRead, OUTLOOK_ATTACH, %INI_FILE%, Outlook, attachments, NONE_VALUE
if OUTLOOK_ATTACH = NONE_VALUE
{
  FileSelectFolder, OUTLOOK_ATTACH, %A_Desktop%,
                  , Select folder to save Outlook attachments
  IniWrite %OUTLOOK_ATTACH%, %INI_FILE%, Outlook, attachments
  if ErrorLevel
  {
    MsgBox Error: Could not write %OUTLOOK_ATTACH% to %INI_FILE%
    Goto, end_hotkey
  }
}
contact_name := get_contact_name_from_outlook_subject()
if contact_name = %NONE_VALUE%
{
    MsgBox,, Save Outlook attachments, No contact name found in e-mail title
    Goto, end_hotkey
}
step_progress_bar()
FileCreateDir, %OUTLOOK_ATTACH%\%contact_name%
; attachment selection window
step_progress_bar()
Send !fna
WinWait,Save All Attachments,,10,
only_one_attachment := 0
if ErrorLevel
{
    ; only one attachment
    only_one_attachment := 1
    Send !fn{enter}
    WinWait,Save Attachment,,0.1,
    if ErrorLevel
        Goto, end_hotkey
}
else
{
    WinActivate
    Send {enter}
    ; file save dialog
    WinWait,Save All Attachments,,0.1,
    if ErrorLevel
        Goto, end_hotkey
}
step_progress_bar()
Send {home}
clipboard = %OUTLOOK_ATTACH%\%contact_name%\
Send ^v{enter}
; exit hotscript if you get a file overwrite warning message
WinWait,Microsoft Office Outlook,,0.1,
if !ErrorLevel
{
    Goto, end_hotkey
}
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + z] Bugzilla search
;----------------------------------------------------------------------
#z::
create_progress_bar("Bugzilla search")
add_progress_step("Reading bug number from selection")
add_progress_step("Opening web link")
copy_to_clipboard()  
step_progress_bar()
if ErrorLevel
  Goto, end_hotkey
; the clipboard takes some time to update
found := RegExMatch(clipboard, "(\d+)", bug)
if found
{
  step_progress_bar()
  Run http://uk00083/show_bug.cgi?id=%bug1%
}
Goto, end_hotkey

end_hotkey_with_error:
end_hotkey:
kill_progress_bar()
Return