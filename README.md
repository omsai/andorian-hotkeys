Make repetitive work awesome!
=============================
<img src="https://cloud.github.com/downloads/omsai/andorian-hotkeys/andorian-scripts-banner.png"
 alt="hot-scripts logo" title="Happy Andorian" align="right" />

Automate your computer to save cumulative hours of your life a year,
and watch your mundane work do itself.

Autohotkey is a powerful scripting language that sends keystrokes,
runs programs and processes information so that you can focus on
important things in life.

Batteries are included: these scripts ask for your information and
save it to an ini file, so no code editing is required to get running.


Usage
-----
#### Global Hotkeys

`⊞ Win`+`W` List all RMAs in web browser

`⊞ Win`+`M` Launch WIP RMA database (automatic login) and optionally opens RMA from highlighted selection

`Ctrl`+`S` (Local hotkey: WIP RMA database) Save WIP record

`Ctrl`+`P` (Local hotkey: WIP RMA database) Open record in Microsoft Word for printing

`⊞ Win`+`T` Open SalesLogix ticket from clipboard or Outlook e-mail title

`⊞ Win`+`O` Find sales order from SO number or serial number that is highlighted or copied

`⊞ Win`+`B` BOM search from highlighted selection

`⊞ Win`+`S` SalesLogix ticket group search

`⊞ Win`+`Z` Bugzilla search from highlighted selection

`⊞ Win`+`5` Paste your initials and today's date

`⊞ Win`+`N` Launch plain text editor


#### USA Hotkeys

`⊞ Win`+`P` Launch US Sales Plan

`⊞ Win`+`I` Fill out BOM descriptions in loan agreement

`Ctrl`+`E` (Local hotkey: Outlook) Archive e-mail in shared inbox


Installation
------------
*Google Chrome* is recommended, since it is used for development and testing.
Please report any issues you may encounter with other browsers.
Mozilla Firefox works but is not ideal since, unlike Chrome and Explorer, it 
prompts for Windows credentials when accessing Intranet resources.

1.  Install [Autohotkey](http://ahkscript.org/)

2.  Install [GitHub for Windows](http://windows.github.com/)

3.  In "GitHub for Windows",
    login to your GitHub account (you probably need to signup for a free
    account first).
    After login enter your e-mail and name for Git.

4.  Login to your GitHub account on this webpage and 
    [clone this repository](github-windows://openRepo/https://github.com/omsai/andorian-hotkeys)

5.  Have the script startup automatically in Windows by
    making a shortcut to `main.ahk` in your Windows start menu > Startup folder

6.  Create the SLX auto-login shortcut by right-clicking on `SalesLogix-Login`
    and pinning it to the Start Menu


Create your own hotkeys
-----------------------
In the same directory as main.ahk, create your ahk file and edit
[main.ahk](andorian-hotkeys/blob/master/main.ahk#L15) to `#Include` it


Contribute
----------
Code and feature requests welcome.

#### Emacs users

You can use `ahk-mode.el` bundled with AutoHotKey.
For the lazy here is an implementation of that file's install instructions:

1.  Add this to your emacs init file, like `~/.emacs`
```
(add-to-list 'load-path "C:/Program Files/AutoHotkey/Extras/Editors/Emacs/")
(setq ahk-syntax-directory "C:/Program Files/AutoHotkey/Extras/Editors/Syntax/")
(add-to-list 'auto-mode-alist '("\\.ahk$" . ahk-mode))
(autoload 'ahk-mode "ahk-mode")
```

2.  Compile `ahk-mode.el` by opening it in Emacs and in the top toolbar clicking:
```
Emacs-Lisp > Byte-compile This File
```
The file is located here
```
C:/Program Files/AutoHotkey//Extras/Editors/Emacs/ahk-mode.el
```

3.  Restart Emacs
