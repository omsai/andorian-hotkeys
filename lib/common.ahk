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
    global SetTitleMatchMode
    SetTitleMatchMode, Slow
    SetTitleMatchMode, 2
    
    _match = From
    _offset = 1
    
    ; Check highlighted Outlook e-mail title for _contact_name
    WinGetText, text, Microsoft Outlook
    ;MsgBox DEBUG: %text%
    _found = 0
    Loop, Parse, text, `n, `r ; parses variable text by newline
    {
        if A_LoopField = %_match%
        {
            _found = 1
        }
        if _found = 1
        {
            if _offset = 0
            {
                ;MsgBox DEBUG: A_LoopField = %A_LoopField%
                return get_legal_filename(A_LoopField)
            }
            else
                _offset := _offset - 1
        }
    }
    
    global NONE_VALUE
    return NONE_VALUE ; No match found
}
