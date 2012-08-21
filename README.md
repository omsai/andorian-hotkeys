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

`⊞ Win`+`A` Save Outlook attachments to folder

`⊞ Win`+`O` Sales order search from highlighted selection

`⊞ Win`+`B` BOM search from highlighted selection

`⊞ Win`+`Z` Bugzilla search from highlighted selection

`⊞ Win`+`5` Paste your initials and today's date


#### USA Hotkeys

`⊞ Win`+`S` Open Shipment form, Shipped folder, and Loan agreements folder

`⊞ Win`+`P` Launch US Sales Plan

`⊞ Win`+`I` Fill out BOM descriptions in loan agreement

`⊞ Win`+`Q` Import folder of multi-tiff data into iQ


Installation
------------
*Google Chrome* is recommended, since it is used for development and testing.
Please report any issues you may encounter with other browsers.

1.  Install [Autohotkey](http://www.autohotkey.com/download/)

2.  Install [GitHub for Windows](http://windows.github.com/)

3.  In "GitHub for Windows",
    login to your GitHub account (you probably need to signup for a free
    account first).
    After login enter your e-mail and name for Git.

4.  Login to your GitHub account on this webpage and 
    [clone this repository](github-windows://openRepo/https://github.com/omsai/andorian-hotkeys)

5.  Have the script startup automatically in Windows by
    making a shortcut to `main.ahk` in your Windows start menu > Startup folder


Create your own hotkeys
-----------------------
In the same directory as main.ahk, create your ahk file and edit
[main.ahk](andorian-hotkeys/blob/master/main.ahk#L15) to `#Include` it


Contribute
----------
Code and feature requests welcome.
