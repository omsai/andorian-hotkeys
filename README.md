Make repetitive work awesome! (RoW Edition)
===========================================
<img src="https://cloud.github.com/downloads/omsai/andorian-hotkeys/andorian-scripts-banner.png"
 alt="hot-scripts logo" title="Happy Andorian" align="right" />

Automate your computer to save cumulative hours of your life a year,
and watch your mundane work do itself.

Autohotkey is a powerful scripting language that sends keystrokes,
runs programs and processes information so that you can focus on
important things in life.

Batteries are included: these scripts ask for your information and
save it to an ini file, so no code editing is required to get running.

NOTE - This is a fork from Pariksheet's (GitHub name Omsai) original work - I've simplified it so that I can maintain it for the RoW Team.


Usage
-----

`⊞ Win`+`?` shows the hotkeys active in this script

`⊞ Win`+`M` Launch WIP RMA database (automatic login) and optionally opens RMA from highlighted selection

`⊞ Win`+`T` Open SalesLogix ticket from clipboard or Outlook e-mail title

`⊞ Win`+`O` Find sales order from SO number or serial number that is highlighted or copied

`⊞ Win`+`B` BOM search from highlighted selection

`⊞ Win`+`Z` Bugzilla search from highlighted selection

`⊞ Win`+`V` Paste your initials and today's date

`⊞ Win`+`N` Opens a text editor - either Emacs, Notepad++ or Notepad

`⊞ Win`+`S` Find ship date from Sales Order

`⊞ Win`+`X` Find SLX Tickets from Sales Order



Installation
------------
*Google Chrome* is recommended, since it is used for development and testing.
Please report any issues you may encounter with other browsers.
Mozilla Firefox works but is not ideal since, unlike Chrome and Explorer, it 
prompts for Windows credentials when accessing Intranet resources.

1.  Install [Autohotkey](http://www.autohotkey.com/download/)

2.  Install [GitHub for Windows](http://windows.github.com/)

3.  Register for a GitHub account [here](https://github.com/join)
	
4.	In "GitHub for Windows",
    login to your GitHub account (you probably need to sign up for a free
    account first).
    After login enter your e-mail and name for Git.

5.  Login to your GitHub account on this webpage and 
    [clone this repository](github-windows://openRepo/https://github.com/JimboMahoney/andorian-hotkeys)

6.  Have the script startup automatically in Windows by
    making a shortcut to `main.ahk` in your Windows start menu > Startup folder

7.  Create the SLX auto-login shortcut by right-clicking on `SalesLogix-Login`
    and pinning it to the Start Menu


Create your own hotkeys
-----------------------
In the same directory as main.ahk, create your ahk file and edit
[main.ahk](andorian-hotkeys/blob/master/main.ahk#L18) to `#Include` it


Contribute
----------
Code and feature requests welcome.


