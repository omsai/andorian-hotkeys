﻿; Boilerplate
#NoEnv ; performance and compatibility
SetWorkingDir %A_ScriptDir%  ; consistent starting directory.

; Libraries
#Include lib\common.ahk
#Include lib\progress_bar.ahk
#Include lib\saleslogix.ahk
; (Optional) Additional shortcuts which would not work for everyone
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
  Run http://andor.andortech.net/wipReports/rma.asp
  WinWait %RMA_BROWSER_TITLE%,,120
  if ErrorLevel
  {
    progress_error(A_LineNumber, "Browser timeout")
    Gosub, end_hotkey_with_error
  }
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
  Gui, Destroy
  Gui, Add, Text,, Select your region:
  Gui, Add, DropDownList, vREGION, AndorUS||AndorUK|AndorJapan|All
  Gui, Add, Button, Default, OK
  Gui, Show
  Return
  
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
    ; If no RMA is displayed, sending ESC would close the Andor (Live)
    ; window.  So only send ESC if the RMA field is not empty.
    ControlGetText, existing_RMA, ThunderRT6TextBox17 ; "RMA No" text.
    if existing_RMA !=
    {
      Send {Escape}
    }
  }

  begin := RegExMatch(clipboard, "R\d{5}")
  if begin != 0
    begin++  ; we don't want the R in the matched string
  else  ; fallback to checking just digits
    begin := RegExMatch(clipboard, "\d{5}")
  if begin != 0
  {
    rma := SubStr(clipboard, begin, 5)
    SetKeyDelay, -1
    Send R%rma%{tab}
    SetKeyDelay, 10		; reset to default value
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
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
WinActivate
SetKeyDelay, -1
Send %RMA_USER%{tab}%RMA_PASS%{tab}{Enter}
SetKeyDelay, 10			; reset to default value

if ErrorLevel
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}

step_progress_bar()
WinActivate
Send {Down}{Enter}{Down}{Down}{Enter}

WinWait, Andor (Live),,10
if ErrorLevel
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}

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

  control_wait_focus(my_control, delay:=100, loops:=40)
  {
    ; Waits till the `my_control` is in focus again.
    Loop, %loops%
    {
      ControlGetFocus, control
      if control = %my_control%
      {
	ErrorLevel = 0
	return
      }
      else
	Sleep, %delay%
    }
    ErrorLevel = 1
  }

  ; Check if on secondary page
  ControlGetText, page, ThunderRT6CommandButton8 ; "Page" button label
  toggle_page = 0
  if page = F7- Prev Page
  {
    toggle_page = 1
  }  
  ControlGetText, clipboard, ThunderRT6TextBox17 ; "RMA No" textbox
  ClipWait, 3
  if clipboard !=
  {
    step_progress_bar()
    Progress,,Saving RMA %clipboard%
    ControlClick, ThunderRT6CommandButton6 ; "CR-Accept" button
    control_wait_focus("ThunderRT6TextBox17") ; "RMA No" text field
    if ErrorLevel
    {
      Goto, End_hotkey_with_error
    }
    step_progress_bar()
    Send ^v{Tab}
  }
  if toggle_page
  {
    control_wait_focus("ThunderRT6TextBox14") ; "Status" text field
    Send {F6}
    control_wait_focus("ThunderRT6TextBox1") ; "Work Carried Out" textbox
    Send ^{End}
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
  
  WinWait,, 100`%, 10
  IfWinNotExist
  {
    Goto, end_hotkey
  }
  step_progress_bar()
  Loop, 10 ; loop since there's no way to know if the button has been clicked
  {
    Sleep, 500
    IfWinExist, Export ahk_exe Andor.exe ahk_class #32770
    {
      WinActivate
      Break
    }
    Else
    {
      ControlClick, X320 Y40 ; Export button is at 326, 41
    }
  }
  
  Sleep, 1000 ; sometimes the CSV window opens instead
  Send {Tab 2}{End}{Tab}{Home}{Enter}
  ; close export and print window
  step_progress_bar()
  WinWaitClose, Exporting Records,, 20
  IfWinExist, Export
  {
    WinActivate
    Send {Esc}
  }
  WinWaitActive, ~Re ahk_class OpusApp,, 10 ; Wait for Word to open
  ; Sleep, 5000 ; otherwise Alt+F4 below doesn't work
  IfWinExist,, 100`%
  {
    WinActivate
    Send !{F4}
  }
#IfWinActive ; turn off context sensitivity
Goto, end_hotkey

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
FormatTime, TimeVar, A_Now, dd-MMM-yyyy
SetKeyDelay, -1
Send %TimeVar% %INITIALS%{Enter}{Enter}
SetKeyDelay, 10			; reset to default value
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + o] Sales order search from clipboard
;----------------------------------------------------------------------
#o::
  create_progress_bar("Sales Order search")
  copy_to_clipboard()
  matches =			; Clear old matches.
  
  ; there's no native function to parse several regex matches, so one has to
  ; reuse the `begin` position parameter to check the full string
  begin = 1
  While begin := RegExMatch(clipboard, "(((CN\d{6})|([MX]?\d{6})|(u\d{5,6})|(R\d{5})|(D\d{4,5}))[\/]?\d?)"
                            , match
                            , begin + StrLen(match))
  {
    sales_order%A_Index% := match1
    matches := A_Index
  }
  Loop, %matches%
  {
    add_progress_step("Querying sales order")
  }
  
  ; open a new window for each match
  Loop, %matches%
  {
    step_progress_bar()
    order := sales_order%A_Index%
    Run http://andor.andortech.net/cm.mccann/Sales Orders/dbSearch.asp?order_no=%order%
  }

  ; If there are no matches for sales orders, assume the selection is
  ; a serial number and find the corresponding sales order(s).
  if matches !=
    Goto, End_hotkey  ; Sales orders were in clipboard, so end hotkey.

  clipboard:=strip(clipboard)	; Remove whitespace, CR, LF, commas, etc.
  add_progress_step("Querying serial# '" . clipboard . "'")
  add_progress_step("Waiting for Enter Values window")
  Run http://intranet2/reports/ViewReport.aspx?ReportPath=I:\Intranet\Reports\Sales+Information\Utilities\Orders+with+serial+no.rpt
  step_progress_bar()
  WinWait, Report Viewer,,30
  If ErrorLevel
  {
    progress_error(A_LineNumber, "Browser timeout")    
    Goto, End_hotkey
  }
  step_progress_bar()
  Sleep, 3000
  WinActivate
  Send {Tab 3}%clipboard%{Enter}
  Goto, End_hotkey

;----------------------------------------------------------------------
; [Windows Key + b] BOM search from clipboard
;----------------------------------------------------------------------
#b::
create_progress_bar("BOM search")
copy_to_clipboard()
clipboard := strip_line_ends(clipboard)
clipboard := Trim(clipboard)	; Part numbers can contain spaces.
add_progress_step("Opening web page")
add_progress_step("Querying '" . clipboard . "'")
step_progress_bar()
IfWinNotExist, Shamrock Components
{
  Run http://andor.andortech.net/bomreport/
  WinWait, Shamrock Components,,40
}
if ErrorLevel
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
step_progress_bar()
WinActivate
Click 10, 105
while (A_Cursor = "AppStarting")
  Sleep,500
; Scroll to beginning of listbox so that previous queries don't affect
; result.
Send {tab}{Home}
; Select part number.
Send %clipboard%{tab}
; Click "Display"
Send {Enter}
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + t] Ticket search from clipboard or Outlook e-mail title
;----------------------------------------------------------------------
#t::
create_progress_bar("Ticket search")
add_progress_step("Extracting Ticket ID")
add_progress_step("Opening ticket in SalesLogix")
add_progress_step("Opening contact in SalesLogix")
copy_to_clipboard()
step_progress_bar()
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
  ; Open contact if no ticket number found, so that a new ticket can
  ; be created.
  step_progress_bar()
  step_progress_bar()

  contact_email := get_email_from_clipboard()

  if contact_email = %NONE_VALUE%
    contact_email := get_contact_email_from_outlook_subject()
  
  if contact_email = %NONE_VALUE%
  {
    MsgBox,, Ticket search, No ticket number found in clipboard or e-mail title
    Goto, end_hotkey
  }
  else
  {
    open_contact_by_email(contact_email)
    Goto, end_hotkey
  }
}
step_progress_bar()
open_ticket(ticket)
Goto, end_hotkey

;----------------------------------------------------------------------
; [Ctrl + x] Ticket Activity classification
;----------------------------------------------------------------------
#IfWinActive Sage SalesLogix
!x::
Gui, Destroy
Gui, Add, Radio, vSENT Group Checked, &Sent E-Mail
Gui, Add, Radio,, &Received E-Mail
Gui, Add, Radio, vCLASS Group Checked, 1-&Customer
Gui, Add, Radio,, 5-&Internal
Gui, Add, Button, Default gapply_classification, OK
Gui, Show
#IfWinActive
Return

apply_classification:
Gui, Submit
WinActivate, Sage SalesLogix
Send {AppsKey}o
If SENT = 1
{
  activity = Sent E-Mail
} Else {
  activity = Received E-Mail
}
If CLASS = 1
{
  access = 1-Customer
} Else {
  access = 5-Internal
}
SetKeyDelay, -1
WinWait, Edit Ticket Activity,, 5
Sleep, 200
Control, EditPaste, %activity%, TEditEx4
Send {Tab}
Sleep, 200
Control, EditPaste, %access%, TEditEx3
Send {Tab}
Sleep, 200
SetKeyDelay, 10			; reset to default value
Send !o
Gui, Destroy
Return

;----------------------------------------------------------------------
; [Windows Key + n] Text Editor
;----------------------------------------------------------------------
#n::
create_progress_bar("Launch Text Editor")
ErrorLevel = ERROR
; Editor preference in descending order
emacs := A_ProgramFilesX86 . "\ErgoEmacs\ErgoEmacs.exe"
editors = %emacs%,notepad++,notepad
Loop, parse, editors, `,
{
  Run, %A_LoopField%, %A_Desktop%, UseErrorLevel
  if ErrorLevel = 0
    Goto, end_hotkey
}

progress_error(A_LineNumber)
Goto, end_hotkey_with_error

;----------------------------------------------------------------------
; [Windows Key + s] Sales order search from clipboard
;----------------------------------------------------------------------
#s::
; Create GUI
Gui, Destroy
Gui, Add, Text, Right c40, Search for text:
Gui, Add, Text, Right, In any of these fields:
Gui, Add, Text,
Gui, Add, Text,
Gui, Add, Text, Right, Using this Ticket group:
Gui, Add, Edit, vSEARCH_TEXT ym,
Gui, Add, Radio, vTYPE Group, &Subject or Description or Resolution
Gui, Add, Radio, Checked, Sales &Order
Gui, Add, ComboBox, vGROUP_NAME, _Search Ticket|System Tickets|USA System Installations||
Gui, Add, Button, Default xm gapply_group, OK
Gui, Show,, Ticket Group Search
Return

apply_group:
Gui, Submit
step_progress_bar()
; Create conditions array.
conditions := Object()		
If TYPE = 1
{
  conditions.Insert("Ticket.Subject")
  conditions.Insert("Ticket.TicketProblem.Notes")
  conditions.Insert("Ticket.TicketSolution.Notes")
  andor := "OR"
}
Else 
{
  conditions.Insert("Ticket.Userfield2")
  andor := "AND"
}
Gui, Destroy
SEARCH_TEXT := Trim(SEARCH_TEXT)
create_progress_bar("Ticket group search")
add_progress_step("Filtering '" . SEARCH_TEXT . "'")
step_progress_bar()
copy_group_adding_conditions("Ticket", GROUP_NAME, SEARCH_TEXT, conditions, andor)
kill_progress_bar()
Return

;----------------------------------------------------------------------
; [Windows Key + z] Bugzilla search
;----------------------------------------------------------------------
#z::
create_progress_bar("Bugzilla search")
add_progress_step("Reading bug number from selection")
add_progress_step("Opening web link")
copy_to_clipboard()  
step_progress_bar()
found := RegExMatch(clipboard, "(\d+)", bug)
if found
{
  step_progress_bar()
  Run http://be-qa-01/show_bug.cgi?id=%bug1%
}
Goto, end_hotkey

;----------------------------------------------------------------------
; [Windows Key + /] Help
;----------------------------------------------------------------------
#/::
Gui, Destroy
FileRead, readme, README.md
If ErrorLevel
{
  Gui, Add, Text,,Error: unable to open README file
}
Else
{
  begin := RegexMatch(readme, "Usage\r*\n*---[-]*")
  end := RegexMatch(readme, "\r*\n*Installation\r*\n*---[-]*", "", begin + 20)
  shortcuts := SubStr(readme, begin, end - begin)
  ; Remove Windows characters since they don't render properly.
  new_shortcuts := RegexReplace(shortcuts, ".*Win", "``Win")
  Gui, Add, Text,,%new_shortcuts%
}
Gui, Show,, Keyboard Shortcuts
;WinSet, Transparent, 150
Return


;; Begin USA hotkeys
RegRead, time_zone, HKEY_LOCAL_MACHINE
  , System\CurrentControlSet\Control\TimeZoneInformation, Standard
If %time_zone% = "Eastern Standard Time"
{
;----------------------------------------------------------------------
; [Windows Key + p] Launch US Sales Plan
;----------------------------------------------------------------------
#p::
create_progress_bar("Launch US Sales Plan")
Run, excel.exe /r "\\ct-dc-01\home\Common\Sales Plan\Orders Query - Open & Closed.xls"
kill_progress_bar()
Return

;----------------------------------------------------------------------
; [Windows Key + i] Fill out BOM descriptions in loan agreement
;----------------------------------------------------------------------
#i::
create_progress_bar("Fill Loan Agreement")
; clear variables
description = 
WinGet, word_window_id, ID, A
MyLabel:
copy_to_clipboard()
clipboard := strip_line_ends(clipboard)
clipboard := Trim(clipboard)	; Part numbers can contain spaces.
; break repeat loop if nothing copied
if clipboard =
{
  kill_progress_bar()
  Return
}
Progress,,Lookup "%clipboard%"

; focus on BOM window
IfWinNotExist, Shamrock Components
  Run http://andor.andortech.net/bomreport/
WinWait, Shamrock Components,,20
if ErrorLevel
{
  kill_progress_bar()
  Return
}
WinActivate

; focus on product code
Click 10, 105
while (A_Cursor = "AppStarting")
  Sleep,500
Sleep, 1000
focus_on_browser_page()
Click 10, 105			; fix for Chromium issue# 181144
Send {tab}

; Display detail view
Send {Home}%clipboard%{tab}{Enter}
Sleep,1000
Click 10, 105
while (A_Cursor = "AppStarting")
  Sleep,500
Sleep,1000

; Copy description
Click 10, 105			; focus on browser page, not iframe.
Send {tab 3}^a			; 3 tab keystrokes to get to iframe.
copy_to_clipboard()
end := InStr(clipboard, "[Print]")
if end != 0
{
    StringLeft, description, clipboard, end-1
    clipboard = %description%
    ; Remove all CR+LF's from the clipboard contents:
    StringReplace, clipboard, clipboard, `r`n, , All
    WinActivate, ahk_id %word_window_id%
    Send {tab}^v
}
else
{
    MsgBox InStr failed to find [Print]
    kill_progress_bar()
    Return
}
Send {tab 2} ; focus on next Word table field
Goto, MyLabel

;----------------------------------------------------------------------
; [Ctrl + e] Archive Shared Inbox E-mail to "Old Inbox" folder
;----------------------------------------------------------------------
#IfWinActive us_microscopy_support - Inbox in Mailbox - us_microscopy_support - Microsoft Outlook
^e::
create_progress_bar("Archive E-mails to 'Old Inbox' folder")
add_progress_step("Mark as read")
add_progress_step("Moving to 'Old Inbox' folder")
step_progress_bar()
Send ^q
step_progress_bar()
Send ^+v{Home}Mailbox - us_m{Right}In{Right}Old{Enter}
#IfWinActive
kill_progress_bar()
Return

}  ;; End USA hotkeys

GuiClose:
Gui, Destroy
Return

GuiEscape:
Gui, Destroy
Return

end_hotkey_with_error:
end_hotkey:
kill_progress_bar()
Return
