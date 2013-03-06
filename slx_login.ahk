; Login to SLX

#Include lib\common.ahk ; for A_ProgramFilesX86

IniRead, SLX_PASS, %INI_FILE%, SLX, password, NONE_VALUE
if SLX_PASS = NONE_VALUE
{
  InputBox SLX_PASS, New SLX user, Enter SLX password, HIDE
  IniWrite %SLX_PASS%, %INI_FILE%, SLX, password
  if ErrorLevel
  {
    MsgBox Error: Could not write your password to %INI_FILE%
    Exit, 1
  }
}

Run %A_ProgramFilesX86%\SalesLogix\SalesLogix.exe
WinWait, Please log on...  ; startup can take several minutes, so no timeout
Sleep, 500
WinActivate
Send %SLX_PASS%{Enter}
