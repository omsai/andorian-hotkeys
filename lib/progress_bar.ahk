; Show user script progress

create_progress_bar(hotkey_title)
{
  global
  steps = 0
  current_step = 0
  Progress, M Y50, Running..., %hotkey_title%, HotKey Progress
}

kill_progress_bar()
{
  Progress, 100, Finished
  Sleep, 500
  Progress, Off
}

add_progress_step(step_description)
{
  global
  steps += 1
  step%steps% := step_description
}

step_progress_bar()
{
  global
  current_step += 1
  local _bar_position := 100 * current_step /(steps + 1)
  local _step_description := step%current_step%
  Progress, %_bar_position%, %_step_description%...
}