Make repetitive work awesome!
=============================
<img src="https://github.com/downloads/omsai/andorian-hotkeys/andorian-scripts-banner.png"
 alt="hot-scripts logo" title="Happy Andorian" align="right" />

Automate your computer to save cummulative hours of your life a year,
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

`⊞ Win`+`T` Open SalesLogix ticket from clipboard or Outlook e-mail title

`⊞ Win`+`A` Save Outlook attachments to Desktop ticket folder

`⊞ Win`+`O` Sales order search from highlighted selection

`⊞ Win`+`B` BOM search from highlighted selection

`⊞ Win`+`D` Paste your initials and today's date


#### USA Hotkeys

`⊞ Win`+`S` Open Shipment form, Shipped folder, and Loan agreements folder

`⊞ Win`+`P` Launch US Sales Plan

`⊞ Win`+`I` Fill out BOM descriptions in loan agreement

`⊞ Win`+`Q` Import folder of multi-tiff data into iQ


Installation
------------
1.  Install [Autohotkey](http://www.autohotkey.com/download/)

2.  Install [Git](http://help.github.com/win-set-up-git/)

3.  Clone this existing repository
    *  Navigate to the directory you want to install this folder.  I like to use  target `%HOMEPATH%`
    *  `(Right-click) in the folder > Git Bash`
    *  Type `git clone https://omsai@github.com/omsai/andorian-hotkeys.git`
    *  When prompted for a password just hit enter

4.  Have the script startup automatically with Windows by
    making a shortcut to `main.ahk` in your Windows start menu > Startup folder

5.  (Optional) Create your own hotkeys.
    Create your ahk file and `#Include` it in [main.ahk](andorian-hotkeys/blob/master/main.ahk#L15)


Contribute
----------
Code and feature requests welcome.