; Configuration saving
INI_FILE := "ahk.ini"  ; Persistent script variables
NONE_VALUE := "NONE"   ; No script variable can be this value

; Windows 7 compatibility
if %A_Is64bitOS%
{
  A_ProgramFilesX86 := A_ProgramFiles . " (x86)"
}
else {
  A_ProgramFilesX86 := A_ProgramFiles
}

; Name all function variables with leading underscore
; to avoid local variable error messages


copy_to_clipboard()
{
  clipboard = ; empty the clipboard
  Send ^c
  ClipWait, 2
}

focus_on_browser_page()
{
  ; Chrome
  Send ^l{tab}
}

get_legal_filename(string) {
  return RegExReplace(string, "[<>:\/\\\|\?\*]", "_")
}

get_contact_name_from_outlook_subject()
{
  _MailItems := ComObjActive("Outlook.Application").ActiveExplorer.Selection
  _MailItem := _MailItems.Item(1)
  _contact_name := _MailItem.SenderName
  if (_contact_name = "")
  {
    global NONE_VALUE
    return NONE_VALUE ; No match found
  }
  else
    return _contact_name
}

get_contact_email_from_outlook_subject()
{
  _MailItems := ComObjActive("Outlook.Application").ActiveExplorer.Selection
  _MailItem := _MailItems.Item(1)
  _contact_email := _MailItem.SenderEmailAddress
  if (_contact_email = "")
  {
    global NONE_VALUE
    return NONE_VALUE ; No match found
  }
  else
    return _contact_email
}

get_ticket_number_from_outlook_subject()
{
  ; Check clipboard
  _begin := RegExMatch(clipboard, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
  if (_begin != 0)
  {
    _ticket := SubStr(clipboard, _begin, 13)
    ; Prevent this clipboard data from clobbering subsequent Outlook subject
    ; selections
    clipboard = ; clear clipboard
    return _ticket
  }
    
  ; Check highlighted Outlook e-mail title for _ticket number
  _MailItems := ComObjActive("Outlook.Application").ActiveExplorer.Selection
  _MailItem := _MailItems.Item(1)
  _subject := _MailItem.Subject
  _begin := RegExMatch(_subject, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
  if (_begin != 0)
  {
    _ticket := SubStr(_subject, _begin, 13)
    ;MsgBox %_ticket%
    return _ticket
  }

  global NONE_VALUE
  return NONE_VALUE ; No match found
}
