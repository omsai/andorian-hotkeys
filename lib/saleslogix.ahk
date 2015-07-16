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

open_systemticket()
{
    If !_get_saleslogix_window()
      return
    WinMenuSelectItem,,,Lookup,Tickets,Advanced Lookup
	SetKeyDelay, -1
	WinWait, Advanced Lookup,,10
    if ErrorLevel
        return
    WinActivate
    Send {tab}
	Send %clipboard%
	Sleep 500
	SetKeyDelay, 10
	Send {tab}{tab}{tab}
	SetKeyDelay, -1
	Send SERVICE
	Sleep 100
	SetKeyDelay, 10
	Send {tab}
	SetKeyDelay, -1
	Send CONFOCAL
	Sleep 100
	SetKeyDelay, 10
	Send {tab}
	SetKeyDelay, -1
	Send INSTALLATION
	SetKeyDelay, 10
	WinWait, Advanced Lookup,,10
    if ErrorLevel
        return
    WinActivate
    Send !o
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
    If _action = copy
    {
      Send +{F10}		; Open Ticket group
      global SetTitleMatchMode
      SetTitleMatchMode, 1
      SetTitleMatchMode, Fast
      WinWait,Sage SalesLogix - [Ticket:,,5
    }
    WinMenuSelectItem,,,View,Groups
    Send %_category%{right}%_name%
    If _action = copy
    {
      Send !c
      WinWait, Query Builder - ,,10
      if ErrorLevel
        Return 0
      WinClose, Group Manager
      WinActivate
    }
    Else
      Send {Enter}
    Return 1
}

copy_group_adding_conditions(_category, _name, _text, _conditions, _andor:="OR")
{
    If !_open_group(_category, _name, "copy")
      return
    _new_name = tmp_%_text%
    Send %_new_name%		; Change the group name.
    ControlClick, TPageControl1	; Properties tab.
    Send {Right}		; Conditions tab.
    for _index, _condition in _conditions
    {
      WinActivate, Query Builder - ,,10
      StringSplit, _arr, _condition, `.
      ; The canonical _arr.MaxIndex() returns nothing, so we have to
      ; manually count the number of items.
      _count := 0
      Loop, %_arr0%
      {
      	_item := _arr%A_Index%
	_count++
      }
      ; Focus on the top of the tree view.
      Click 24, 62
      Send {Home}
      ; Move within the tree.
      If _count = 3
      {
	Send {Right}
	while (A_Cursor = "AppStarting")
	  Sleep, 500
	Send %_arr2%
      }
      ; Focus on field.
      WinGetActiveStats, _z, _w, _h, _x1, _y1
      _x2 := _w - 170		; 157 = width of right button panel.
      _y2 := _h - 304		; 304 = height of bottom tab panel.
      Click %_x2%, 62
      _item := _arr%_count%
      Send %_item%
      Sleep, 500	    	; Wait for h-scroll movement to stop.
      ; Move mouse to the highlighted text using PixelSearch.
      _highlight_color := 0xff9933
      PixelSearch, _x, _y, %_x1%, %_y1%, %_x2%, %_y2%, %_highlight_color%, , Fast
      MouseMove, %_x%, %_y%
      ; Open the Assign Condition window.
      Click,,2
      WinWait, Assign Condition
      Send {Tab}%_text%
      Control, Uncheck,, TCheckBox2 ; Try to uncheck Case Sensitive
      Control, Choose, 2, TComboBox1 ; Select "contains".
      ; ControlFocus, TButton3	; Focus on OK.
      Send {Enter}		; Click OK.
      WinWaitClose, Assign Condition
      If _andor = OR
      {
        ; Set And/Or to OR
        WinActivate, Query Builder - ,,10
	Sleep, 200
	_x1 := 20
	_y1 := _h - 164		; 164 = height of bottom listbox.
	PixelSearch, _x, _y, %_x1%, %_y1%, %_w%, %_h%, %_highlight_color%, , Fast
	_y := _y - 15
	MouseMove, %_x%, %_y%
	Click right
	Sleep, 200
	Send {Up}{Right}{Down}{Enter}
      }
    }
    WinActivate, Query Builder - ,,10
    Send {Enter}
    ; Show result.
    _open_group(_category, _new_name)
}