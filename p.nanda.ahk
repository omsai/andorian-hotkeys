; Shortcuts implemented in this script
;

;----------------------------------------------------------------------
; [Windows Key + f] (Local hotkey) Fill Chrome decontamination form
;----------------------------------------------------------------------
#IfWinActive Andor > decontamination
#f::
  SetKeyDelay, -1
  focus_on_browser_page()
  Send Pariksheet Nanda{Tab}
  Send Andor Technology{Tab}
  Send 8602909211{Tab}
  Send p.nanda@andor.com{Tab}
  Send Product Support Engineer{Tab}
  FormatTime, date, A_Now, M/dd/yyyy
  Send %date%{Tab 5}
  Loop, 4
  {
    Send {Down}{Tab}
  }
  Send {Tab}{Space}
  SetKeyDelay, 10		; reset to default value
  Return
#IfWinActive ; turn off context sensitivity
