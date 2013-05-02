; Boilerplate
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
  Run http://intranet/scripts/wip/rma.asp
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
  Gui, Add, Text,, Select your region:
  Gui, Add, DropDownList, vREGION, AndorUS||AndorUK|AndorJapan|All
  Gui, Add, Button, Default, OK
  Gui, Show
  
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
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
WinActivate
Send %RMA_USER%{tab}%RMA_PASS%{tab}{Enter}

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
  ControlGetText, clipboard, ThunderRT6TextBox17 ; "RMA No" textbox
  if clipboard !=
  {
    step_progress_bar()
    ControlClick, ThunderRT6CommandButton6 ; "CR-Accept" button

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
  
  WinWait,, 100`%, 10
  IfWinNotExist
  {
    Goto, end_hotkey
  }
  step_progress_bar()
  Loop, 10 ; loop since there's no way to know if the button has been clicked
  {
    Sleep, 500
    IfWinExist, Export
    {
      WinActivate
      Break
    }
    Else
    {
      ControlClick, X320 Y40 ; Export button is at 326, 41
    }
  }
  
  Sleep, 500 ; sometimes the CSV window opens instead
  Send {Tab 2}{End}{Tab}{Home}{Enter}
  ; close export and print window
  step_progress_bar()
  WinWaitClose, Exporting Records,, 20
  IfWinExist, Export
  {
    WinActivate
    Send {Esc}
  }
  Sleep, 5000 ; otherwise Alt+F4 below doesn't work
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
  While begin := RegExMatch(clipboard, "((([MX]?\d{6})|(u\d{5,6}))[\/]?\d?)"
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
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
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
WinWait, Edit Ticket Activity,, 5
Clipboard := activity
Send ^v
Sleep, 200
Clipboard := access
Send {Tab}^v{Tab}
Sleep, 200
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
WinWait,Save All Attachments,,1
Send {enter}
WinWait,Save Attachment,,0.1
step_progress_bar()
Send {home}
clipboard = %OUTLOOK_ATTACH%\%contact_name%\
Send ^v{enter}
; force overwrite on seeing a file overwrite warning message
overwrite = 1;
while overwrite
{
  WinWait,Microsoft Office Outlook,,2
  If ErrorLevel
  {
    overwrite = 0
  }
  Else
  {
    WinActivate
    Send !y
  }
}
Run %clipboard%
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
Gui +AlwaysOnTop +ToolWindow
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
  Gui, Add, Text,,%shortcuts%
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
; [Windows Key + s] Open Shipment form, Shipped and Loans folders
;----------------------------------------------------------------------
#s::
create_progress_bar("Shipping form and loan folder")
TimeVar := A_Now
FormatTime, TimeVar, A_Now, yyyy
Run \\ct-dc-01\home\Man Pack List & Loan Agreements\%TimeVar% LOAN AGREEMENTS
Run \\be-fp-01\msystems\ISO 9001 - Quality\FORMS\FM US Shipment Request Form.doc
kill_progress_bar()
Return

;----------------------------------------------------------------------
; [Windows Key + p] Launch US Sales Plan
;----------------------------------------------------------------------
#p::
create_progress_bar("Launch US Sales Plan")
Run, \\ct-dc-01\home\Common\Sales Plan\U.S. Sales Plan.xls
SetTitleMatchMode Regex
WinWait, (Password|File in Use),, 20 ; or file in use
If ErrorLevel
{
  ; Kludge: cannot use Goto, Label in the USA hotkeys block
  kill_progress_bar()
  Return
}
Send !r
kill_progress_bar()
Return

;----------------------------------------------------------------------
; [Windows Key + i] Fill out BOM descriptions in loan agreement
;----------------------------------------------------------------------
#i::
create_progress_bar("Fill Loan Agreement")
; clear variables
description = 
MyLabel:
copy_to_clipboard()
; break repeat loop if nothing copied
if clipboard =
{
  kill_progress_bar()
  Return
}

; focus on BOM window
IfWinNotExist, Shamrock Components
  Run http://intranet/bomreport/
WinWait, Shamrock Components,,20
if ErrorLevel
{
  kill_progress_bar()
  Return
}
WinActivate

; focus on product code
while (A_Cursor = "AppStarting")
    continue
Sleep,400
focus_on_browser_page()
; Send {tab}			; commented out since it breaks in
				; Google Chrome 26.0.1410.64 m
; Display detail view
Send %clipboard%{tab}{Enter}
Sleep,1000
while (A_Cursor = "AppStarting")
    continue

; Copy description
Send {tab}^a
copy_to_clipboard()
end := InStr(clipboard, "[Print]")
Send {tab 7} ; clear web page element focus
if end != 0
{
    StringLeft, description, clipboard, end-1
    clipboard = %description%
    ; Remove all CR+LF's from the clipboard contents:
    StringReplace, clipboard, clipboard, `r`n, , All
    Send !{tab}
    Sleep,100 ; Wait for Word to be active
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

}  ;; End USA hotkeys


GuiEscape:
Gui, Destroy
Return

end_hotkey_with_error:
end_hotkey:
kill_progress_bar()
Return
