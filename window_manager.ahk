; Backport of Windws 7 hotkeys for XP users

/*
Must of the logic in this script has been adapted from Jeremy's blog:
http://cagedechoes.blogspot.com/2011/03/windows-snap-hotkeys-for-xp-using.html  
*/

get_width(left, right)
{
  If (left < 0)
  {
    Return Ceil((Abs(right) - left) / 2)
  }
  Else
  {
    Return Ceil((right - left) / 2)
  }
}
get_left(left, right)
{
  If (left < 0)
  {
    Return left + Ceil((Abs(right) - left) / 2)
  }
  else
  {
    Return left + Ceil((right - left) / 2)
  }
}


#up:: ; Maximize or restore window
  WinGetTitle, title, A
  WinGet, maximized, MinMax, %title%
  if maximized
    WinRestore, %title%    
  else
    WinMaximize, %title%
  Return
  

#down:: ; Minimize
  WinGetTitle, title, A
  WinMinimize, %title%
  Return
  

#left:: ; Snap Left
  WinGetTitle, title, A
  WinRestore, %title%
  SysGet, monitor_, MonitorWorkArea, 1
  l := monitor_Left
  t := monitor_Top
  w := get_width(monitor_Left, monitor_Right)
  h := monitor_Bottom - monitor_Top
  WinMove, %title%,, l, t, w, h
  Return
  
#right:: ; Snap Right
  WinGetTitle, title, A
  WinRestore, %title%
  SysGet, monitor_, MonitorWorkArea, 1
  l := get_left(monitor_Left, monitor_Right)
  t := monitor_Top
  w := get_width(monitor_Left, monitor_Right)
  h := monitor_Bottom - monitor_Top
  WinMove, %title%,, l, t, w, h
  Return
  
