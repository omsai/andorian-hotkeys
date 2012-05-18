get_ticket_number_from_outlook_subject()
{
    global SetTitleMatchMode
    SetTitleMatchMode, Slow
    SetTitleMatchMode, 2
    match_found := 0
    
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
        subject := A_LoopField
        ;MsgBox DEBUG: %subject%
        _begin := RegExMatch(subject, "\d\d\d\-\d\d\-\d\d\d\d\d\d")
        if _begin != 0
        {
            _ticket := SubStr(subject, _begin, 13)
            ;MsgBox %_ticket%
            return _ticket
        }
    }
    
    return 0 ; No match found
}


/* Gets rid of the Modal Window:
   "Sync Client logix has successfully applied transactions to your local
   database.  ...refresh the client now?"
*/
ignore_saleslogix_refresh()
{
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