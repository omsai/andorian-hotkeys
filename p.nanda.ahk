; Shortcuts implemented in this script
;
; [Windows Key + n] Editor
;

;----------------------------------------------------------------------
; [Windows Key + n] Editor
;----------------------------------------------------------------------
#n::
ErrorLevel = ERROR
; Editor preference in descending order
editors = emacs,notepad++,notepad
Loop, parse, editors, `,
{
  Run, %A_LoopField%, %A_Desktop%, UseErrorLevel
  if ErrorLevel = 0
    return
}