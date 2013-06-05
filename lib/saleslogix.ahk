; Name all function variables with leading underscore
; to avoid local variable error messages

_unminimize_saleslogix_window()
{
    ; Returns 1 if SalesLogix window exists and is unminimized or 0 if
    ; the window does not exist.

    ; When SalesLogix is minimized, it's title is "SalesLogix" and
    ; when it's restored it's title is "Sage SalesLogix -".
    WinGet,_min_max,MinMax,SalesLogix
    SetTitleMatchMode, RegEx
    If _min_max = -1
    {
        WinRestore,SalesLogix
	; WinRestore doesn't block execution until Window is
	; restored, so we have to delay till retore happens manually.
	WinGet,_min_max,MinMax,Sage SalesLogix,,(Server)|(Client)
	Return 1
    }
    WinGet,_min_max,MinMax,Sage SalesLogix,,(Server)|(Client)
    If _min_max =
    {
        MsgBox, Error: No SalesLogix window found.  Is it open?
        Return 0
    }
    Return 1
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
    If ! _unminimize_saleslogix_window()
      return
    SetTitleMatchMode, RegEx
    WinWait,SalesLogix,,10,(Server)|(Client)
    if ErrorLevel
        return
    WinActivate
    _ignore_saleslogix_refresh()
    WinMenuSelectItem,,,Lookup,Tickets,Ticket ID
    SetKeyDelay, -1
    Send %_ticket%{tab}{enter}
    SetKeyDelay, 10		; reset to default value
    WinWait, Lookup Ticket,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
}


open_contact_by_email(_email)
{
    If ! _unminimize_saleslogix_window()
      return
    SetTitleMatchMode, RegEx
    WinWait,SalesLogix,,10,(Server)|(Client)
    if ErrorLevel
        return
    WinActivate
    _ignore_saleslogix_refresh()
    WinMenuSelectItem,,,Lookup,Contacts,E-mail
    Send %_email%{tab}{enter}
    WinWait, Lookup Contact,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
}
