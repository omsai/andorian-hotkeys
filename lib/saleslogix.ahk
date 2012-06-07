; Name all function variables with leading underscore
; to avoid local variable error messages


get_ticket_number_from_outlook_subject()
{
    global SetTitleMatchMode
    SetTitleMatchMode, Slow
    SetTitleMatchMode, 2
    
    ; Check clipboard
    _begin := RegExMatch(clipboard, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
    if _begin != 0
    {
        _ticket := SubStr(clipboard, _begin, 13)
        ;MsgBox %_ticket%
        ; Prevent this clipboard data from clobbering subsequent Outlook subject
        ; selections
        clipboard = ; clear clipboard
        return _ticket
    }
    
    ; Check highlighted Outlook e-mail title for _ticket number
    WinGetText, text, Microsoft Outlook
    ;MsgBox DEBUG: %text%
    Loop, Parse, text, `n, `r ; parses variable text by newline
    {
        _subject := A_LoopField
        ;MsgBox DEBUG: %subject%
        _begin := RegExMatch(_subject, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
        if _begin != 0
        {
            _ticket := SubStr(_subject, _begin, 13)
            ;MsgBox %_ticket%
            return _ticket
        }
    }
    
    global NONE_VALUE
    return NONE_VALUE ; No match found
}


_ignore_saleslogix_refresh()
{
    ; Allows refresh from the Modal Window prompt:
    ;
    ;     "Sync Client logix has successfully applied transactions 
    ;      to your local database.  ...refresh the client now?"
    ;
    ; One has to allow refresh, since suppressing the message could lead to
    ; operation against supplanted records
    
    global SetTitleMatchMode
    WinWait, Confirm,,0.1
    
    ; Ignore ErrorLevel from WinWait ... No prompt Window, so nothing to do
    if !ErrorLevel
    {
        WinActivate
        Send !y
        ; Kludge: wait for SLX refresh to complete before finding _ticket.
        ;         Otherwise sending keystrokes while the main toolbar reloads
        ;         creates an access violation and SLX has to be restarted.
        SetTitleMatchMode, 1
        SetTitleMatchMode, Fast
        _counter = 10
        WinWait,Sage SalesLogix -,,5
        while ( ErrorLevel &&  _counter > 0 )
        {
            WinWait,Sage SalesLogix -,,2
            _counter--
        }
        Sleep,4000
    }
}


open_ticket(_ticket)
{
    WinWait,SalesLogix,,10,Personal
    if ErrorLevel
        return
    WinActivate
    _ignore_saleslogix_refresh()
    Send !ltt%_ticket%{tab}{enter}
    WinWait, Lookup Ticket,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
    ; wait for ticket to be actually open
    WinWait,Sage SalesLogix - [Ticket: %_ticket%],,10
    if ErrorLevel
        return
}


add_attachment(_file)
{
    _ignore_saleslogix_refresh()
    Send !if
    clipboard = %_file%
    Send ^v!o
}