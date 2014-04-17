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
        Progress,,Waiting for refresh to complete
        WinActivate
        Send !y
        ; One has to wait for SLX refresh to complete before finding
	;  _ticket. Otherwise sending keystrokes while the main
	; toolbar reloads creates an access violation and SLX has to
	; be restarted.
	Sleep,4000
        _counter = 10
	; Wait for the menu bar to reappear
	SetTitleMatchMode, 1
	SetTitleMatchMode, Fast
	WinWait,Sage SalesLogix -,,5
	WinActivate
	WinMenuSelectItem,,,Edit,Copy Link to Clipboard
        while ( ErrorLevel &&  _counter > 0 )
        {
            _counter--
	    WinWait,Sage SalesLogix -,,5
	    WinActivate
	    WinMenuSelectItem,,,Edit,Copy Link to Clipboard
	    Sleep,2000
        }
        Progress,,Opening ticket in SalesLogix
    }
}


_get_saleslogix_window()
{
    ; Returns 1 if SalesLogix window exists and is ready to accept
    ; keyboard input; otherwise returns 0.

    ; Hack to close annoying Personal Web Server window, which due to
    ; it's name steals focus from other SLX windows.
    WinClose,Sage SalesLogix Personal Web Server

    If ! _unminimize_saleslogix_window()
      Return 0
    SetTitleMatchMode, RegEx
    WinWait,SalesLogix,,10,(Server)|(Client)
    if ErrorLevel
      Return 0
    WinActivate
    _ignore_saleslogix_refresh()
    Return 1
}


open_ticket(_ticket)
{
    If !_get_saleslogix_window()
      return
    WinMenuSelectItem,,,Lookup,Tickets,Ticket ID
    SetKeyDelay, -1
    Send %_ticket%
    SetKeyDelay, 10		; reset to default value
    WinWait, Lookup Ticket,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
}


open_contact_by_email(_email)
{
    If !_get_saleslogix_window()
      return
    WinMenuSelectItem,,,Lookup,Contacts,E-mail
    Send %_email%{tab}{enter}
    WinWait, Lookup Contact,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
}


_open_group(_category, _name, _action:="enter")
{
    ; Choose group results from group manager.  If "_action" is set to
    ; "edit", it opens the Query Builder, otherwise 

    If !_get_saleslogix_window()
      Return 0
    ; Workaround for SLX bug: "Edit" button in Group Manager is
    ; disabled when no group windows are open.
    If _action = edit
    {
      Send +{F10}		; Open Ticket group
      global SetTitleMatchMode
      SetTitleMatchMode, 1
      SetTitleMatchMode, Fast
      WinWait,Sage SalesLogix - [Ticket:,,5
    }
    WinMenuSelectItem,,,View,Groups
    Send %_category%{right}%_name%
    If _action = edit
    {
      Send !e
      WinWait, Query Builder - %_name%,,10
      if ErrorLevel
        Return 0
      WinClose, Group Manager
      WinActivate
    }
    Else
      Send {Enter}
    Return 1
}

edit_group_conditions(_category, _name)
{
    If !_open_group(_category, _name, "edit")
      return
    ControlClick, TPageControl1	   ; Properties tab.
    Send {Right}		   ; Conditions tab.
    ; Wait for the Conditions tab to load.
    Sleep, 100
    ControlGetPos,_x,_y,,,TSLGrid1 ; Fields table.
    _x := _x + 10		   ; Row 1 x-offset.
    _y := _y + 27		   ; Row 1 y-offset.
    ControlClick, X%_x% Y%_y%	   ; Condition row 1.
    ; Set all field values to clipboard.
    _last_field =
    Loop, 5
    {
        Send !e
	WinWait, Assign Condition
        ControlGetText, _field, TDataEdit1
	; Check if we have reached the end.
	If _field = %_last_field%
	    Break
	Else
	   _last_field := _field
	Send {Tab}^v{Enter}
	WinWaitClose, Assign Condition
	Send {Down}
    }
    WinClose, Assign Condition
    WinActivate, Query Builder - %_name%,,10
    Send {Enter}
    ; Show result.
    _open_group(_category, _name)
}
