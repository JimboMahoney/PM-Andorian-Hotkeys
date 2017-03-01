Make repetitive work awesome! (PM Edition)
===========================================
<img src="https://cloud.github.com/downloads/omsai/andorian-hotkeys/andorian-scripts-banner.png"
 alt="hot-scripts logo" title="Happy Andorian" align="right" />

Automate your computer to save cumulative hours of your life a year,
and watch your mundane work do itself.

Autohotkey is a powerful scripting language that sends keystrokes,
runs programs and processes information for repetitive tasks (for example, looking up SOs; Tickets and RMAs).

Batteries are (mostly) included: these scripts ask for your information and
save it to an ini file, so no code editing is required to get running.

NOTE - This is a fork from Pariksheet's (GitHub name Omsai) original work - I've simplified it so that I can maintain it for the PM Team.

LIMITATIONS - Due to different users having different layouts for WIP Tracking, Win + M may require user-specific customisation.


Usage
-----------

`⊞ Win`+`?` shows the hotkeys active in this script

`⊞ Win`+`B` Open BoMs & CoGs (Bill of Materials / Cost of Goods) for the part code (highlighted selection or copied to clipboard)

`⊞ Win`+`I` Open SalesLogix System / Installation ticket from SLX Account (copied to clipboard)

`⊞ Win`+ SHIFT + `I` Open SalesLogix Photostimulation Installation ticket from SLX Account (copied to clipboard)

`⊞ Win`+`M` Launch WIP RMA database (automatic login) and optionally opens RMA (copied to clipboard or highlighted selection)

`⊞ Win`+`N` Opens a text editor - either Emacs, Notepad++ or Notepad

`⊞ Win`+`O` Opens Sales Order from SO number or serial number (highlighted selection or copied to clipboard)

`⊞ Win`+`P` Opens PriceList Update from the part code (highlighted selection or copied to clipboard)

`⊞ Win`+`Q` Find price quoted to customer and/or ship date from Sales Order (highlighted selection or copied to clipboard)

`⊞ Win`+`S` Opens Stock Requests from the part code (highlighted selection or copied to clipboard)

`⊞ Win`+`T` Open SalesLogix ticket from clipboard or Outlook e-mail title

`⊞ Win`+`V` Paste an SLX comment template for updating Tickets (Name, Date, Summary, Original E-mail)

`⊞ Win`+`W` "Where Used" - Finds Level 10 codes containing the copied or highlighted Level 1 code

`⊞ Win`+`Z` Bugzilla search (highlighted selection or copied to clipboard)

Suggestions should be posted as issues here => https://github.com/JimboMahoney/PM-Andorian-Hotkeys/issues



Installation
------------
*Google Chrome* is recommended, since it is used for development and testing.


1.  Install <a target="_blank" href="http://ahkscript.org/" >Autohotkey</a> (Big blue Download button on top right).

2.  Install <a href="http://windows.github.com/" target="_blank">GitHub for Windows</a>

3.  Register for a GitHub account <a href="https://github.com/join" target="_blank">here</a>
	
4.	Run GitHub for Windows (in the Start Menu) and log in to your GitHub account. Patience is required here, as it takes a long time (30 seconds?) to load...
    After login, enter your e-mail and name for Git.

5.  Login to your GitHub account on <a href="https://github.com/JimboMahoney/PM-Andorian-Hotkeys" target="_blank">this</a> webpage and 
    [clone or download](github-windows://openRepo/https://github.com/JimboMahoney/PM-Andorian-Hotkeys), selecting "Open in Desktop" and choosing a sensible directory on your machine and remembering where you save it!<br>
	<img src="https://cloud.githubusercontent.com/assets/7777844/20211125/f87db180-a7f4-11e6-8885-f5ec402212ee.png"
 alt="download" title="clone" align="center" /> <br> 

6.  Have the script startup automatically in Windows by
    making a shortcut to `main.ahk` (in the GitHub, Andorian Hotkeys directory) in your Windows start menu > Startup folder. This is probably easiest by RIGHT-clicking and dragging the main.ahk into the Start Menu, All Programs, Startup. When you let go of the right mouse button, select "Create Shortcut Here"


Create your own hotkeys
-----------------------
1) Edit the main.AHK file (at the risk of breaking existing hotkeys if not done properly)

Or

2) In the same directory as main.ahk, create your ahk file and edit
[main.ahk](PM-Andorian-Hotkeys/blob/master/main.ahk#L18) to `#Include` it


Contribute
----------
<a href="https://github.com/JimboMahoney/PM-Andorian-Hotkeys/issues" target="_blank">Code and feature requests welcome!</a>


