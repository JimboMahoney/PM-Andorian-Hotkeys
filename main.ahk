; See README.md for shortcuts in this script

;------------
; Boilerplate
;------------
#NoEnv ; performance and compatibility
SetWorkingDir %A_ScriptDir%  ; consistent starting directory.


; Libraries
#Include lib\common.ahk
#Include lib\progress_bar.ahk
#Include lib\saleslogix.ahk

;------------------------
; [Windows Key + /?] Help
;------------------------
#/::
Gui, Destroy
FileRead, readme, README.md
If ErrorLevel
{
  Gui, Add, Text,,Error: unable to open README file
}
Else
{
  begin := RegexMatch(readme, "Usage\r*\n*---[-]*")
  end := RegexMatch(readme, "\r*\n*Installation\r*\n*---[-]*", "", begin + 20)
  shortcuts := SubStr(readme, begin, end - begin)
  ; Remove Windows characters since they don't render properly.
  new_shortcuts := RegexReplace(shortcuts, ".*Win", "``Win")
  Gui, Add, Text,,%new_shortcuts%
}
Gui, Show,, Keyboard Shortcuts
;WinSet, Transparent, 150
Return



;-----------------------------------------------------------------------------
; [Windows Key + b] BoMs & CoGs search from clipboard or highlighted selection
;-----------------------------------------------------------------------------
#b::
create_progress_bar("BOM search")
add_progress_step("Opening web page")
add_progress_step("Querying part number")
copy_to_clipboard()
clipboard := RegexReplace(clipboard, "[[:blank:]]") ; remove tabs and spaces
step_progress_bar()
Run http://andor.andortech.net/cm.mccann/BOM and COGS/


WinWait, Shamrock Components,,100
if ErrorLevel
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
WinActivate
;while (A_Cursor = "AppStarting")
  Sleep,500
step_progress_bar()
SetKeyDelay 100
Send {tab 4}%clipboard%{tab}{Enter}
Goto, end_hotkey

;-----------------------------------------------------------------------
; [Windows Key + i] Install Ticket search from SLX Account in clipboard 
;-----------------------------------------------------------------------
#i::
create_progress_bar("System Ticket search")
add_progress_step("Searching for System Ticket...")
;copy_to_clipboard()
step_progress_bar()
open_systemticket()
Goto, end_hotkey




;---------------------------------------------------------------
; [Windows Key + m] Login to RMA WIP and open RMA from clipboard
;---------------------------------------------------------------
#m::

; FIXME: Login credentials are assumed to be  correct.  If login details
;        are wrong the user would need to edit the INI file.
IniRead, RMA_USER, %INI_FILE%, RMA, username, NONE_VALUE
if RMA_USER = NONE_VALUE
{
  InputBox RMA_USER, New RMA user, Enter WIP Tracking username
  IniWrite %RMA_USER%, %INI_FILE%, RMA, username
  if ErrorLevel
  {
    MsgBox Error: Could not write %RMA_USER% to %INI_FILE%
    Goto, end_hotkey
  }
}
IniRead, RMA_PASS, %INI_FILE%, RMA, password, NONE_VALUE
if RMA_PASS = NONE_VALUE
{
  InputBox RMA_PASS, New RMA user, Enter WIP Tracking password, HIDE
  IniWrite %RMA_PASS%, %INI_FILE%, RMA, password
  if ErrorLevel
  {
    MsgBox Error: Could not write your password to %INI_FILE%
    Goto, end_hotkey
  }
}
IniRead, RMA_MENU, %INI_FILE%, RMA, menu, NONE_VALUE
if RMA_MENU = NONE_VALUE
{
  InputBox RMA_MENU, New RMA user, Enter the location of the RMA Process menu in your WIP Tracking (numbers 3 or 4 or 5) - see: www.github.com/JimboMahoney/RoW-Andorian-Hotkeys/issues/1 , ,600
  IniWrite %RMA_MENU%, %INI_FILE%, RMA, menu
  if ErrorLevel
  {
    MsgBox Error: Could not write your menu location to %INI_FILE%
    Goto, end_hotkey
  }
}

Run C:\vb6\Andor\Andor.exe
WinWait, Login,,40
if ErrorLevel
{
  return
}
else
{
  WinActivate
  Send %RMA_USER%{tab}%RMA_PASS%{tab}{Enter}
  
  WinWait, Andor Technology,,500
  if ErrorLevel
  {
    return
  }
  else 
  IniRead, RMA_MENU, %INI_FILE%, RMA, menu
  if RMA_MENU = 5
  {
    WinActivate
    Sleep 100
	SetKeyDelay, 50
    Send {Down}{Enter}{Down}{Enter}{Down}{Enter}
    Sleep 100
	StringReplace, clipboard, clipboard,MA ;Change/remove "MA" from e.g. RMA12345 => R12345
	StringReplace, clipboard, clipboard,[ ;Remove leading [ sometimes found in SLX RMA references
	StringReplace, clipboard, clipboard,] ;Remove trailing ] sometimes found in SLX RMA references
	Send %clipboard%{Enter}
    
    WinWait, Andor (Live),,10
    if ErrorLevel
    {
      return
    }
    else
    {
      WinActivate, Andor (Live)
      return
    }
  }
else if RMA_MENU = 4
  {
    WinActivate
    Sleep 100
	SetKeyDelay, 50
    Send {Down}{Enter}{Down 2}{Enter}
    Sleep 100
	StringReplace, clipboard, clipboard,MA ;Change/remove "MA" from e.g. RMA12345 => R12345
	StringReplace, clipboard, clipboard,[ ;Remove leading [ sometimes found in SLX RMA references
	StringReplace, clipboard, clipboard,] ;Remove trailing ] sometimes found in SLX RMA references
	Send %clipboard%{Enter}
    
    WinWait, Andor (Live),,10
    if ErrorLevel
    {
      return
    }
    else
    {
      WinActivate, Andor (Live)
      return
    }
  }
else
 {
 MsgBox Error: Could not locate RMA Process menu - check the menu entry in ahk.ini or report an issue on the RoW-Andorian-Hotkeys GitHub.
 Goto, end_hotkey
 }
}
SetKeyDelay, -1
Goto, End_hotkey

;------------------------------
; [Windows Key + n] Text Editor
;------------------------------
#n::
create_progress_bar("Launch Text Editor")
ErrorLevel = ERROR
; Editor preference in descending order
emacs := A_ProgramFilesX86 . "\ErgoEmacs\ErgoEmacs.exe"
editors = %emacs%,notepad++,notepad
Loop, parse, editors, `,
{
  Run, %A_LoopField%, %A_Desktop%, UseErrorLevel
  if ErrorLevel = 0
    Goto, end_hotkey
}

progress_error(A_LineNumber)
Goto, end_hotkey_with_error


;---------------------------------------------------------------------------------
; [Windows Key + o] Sales order or serial number search from highlighted selection
;---------------------------------------------------------------------------------
#o::
  create_progress_bar("Sales Order search")
  copy_to_clipboard()
  matches =			; Clear old matches.
  
  ; there's no native function to parse several regex matches, so one has to
  ; reuse the `begin` position parameter to check the full string
  begin = 1
  While begin := RegExMatch(clipboard, "(((CN\d{6})|([M]?\d{6,7})|([m]?\d{6,7})|(u\d{5,6})|(R\d{5})|(DS\d{4,5})|(D\d{4,5})|(X\d{4,5}))[\/]?\d?)"
                            , match
                            , begin + StrLen(match))
  {
    sales_order%A_Index% := match1
    matches := A_Index
  }
  Loop, %matches%
  {
    add_progress_step("Querying sales order")
  }
  
  ; open a new window for each match
  Loop, %matches%
  {
    step_progress_bar()
    order := sales_order%A_Index%
    Run http://andor.andortech.net/cm.mccann/Sales Orders/dbSearch.asp?order_no=%order%
  }

  ; If there are no matches for sales orders, assume the selection is
  ; a serial number and find the corresponding sales order(s).
  if matches !=
    Goto, End_hotkey  ; Sales orders were in clipboard, so end hotkey.

  clipboard:=strip(clipboard)	; Remove whitespace, CR, LF, commas, etc.
  add_progress_step("Querying serial# '" . clipboard . "'")
  add_progress_step("Waiting for Enter Values window")
  Run http://andor.oxinst.com/reports/ViewReport.aspx?ReportPath=I:/Intranet/Reports/Sales+Information/Utilities/Orders+with+serial+no.rpt
  step_progress_bar()
  WinWait, Report Viewer,,30
  If ErrorLevel
  {
    progress_error(A_LineNumber, "Browser timeout")    
    Goto, End_hotkey
  }
  step_progress_bar()
  Sleep, 3000
  WinActivate
  Send {Tab 3}%clipboard%{Enter}
  Goto, End_hotkey
  
;---------------------------------------------------------------------------
; [Windows Key + p] PriceList Update from clipboard or highlighted selection
;---------------------------------------------------------------------------
#p::
create_progress_bar("PriceList update")
add_progress_step("Opening web page")
add_progress_step("Querying part number")
copy_to_clipboard()
clipboard := RegexReplace(clipboard, "[[:blank:]]") ; remove tabs and spaces
step_progress_bar()
Run http://andor.oxinst.com/sageutils/pricelist/index.asp?


WinWait, PriceList Update,,100
if ErrorLevel
{
  progress_error(A_LineNumber)
  Goto, end_hotkey
}
WinActivate
;while (A_Cursor = "AppStarting")
  Sleep,500
step_progress_bar()
SetKeyDelay 10
Send {tab}%clipboard%{tab}{Enter}
Goto, end_hotkey
  
  
;-----------------------------------------------------------------------------------------------
; [Windows Key + q] Price Quoted / Shipping date from SO from clipboard or highlighted selection
;-----------------------------------------------------------------------------------------------
#q::
  create_progress_bar("Price Quoted / Ship date search")
  copy_to_clipboard()
  ;matches =			; Clear old matches.
  
  clipboard:=strip(clipboard)	; Remove whitespace, CR, LF, commas, etc.
  add_progress_step("Querying Sales Order '" . clipboard . "'")
  add_progress_step("Waiting for Enter Values window")
  Run http://andor.andortech.net/reports/ViewReport.aspx?ReportPath=I:/Intranet/Reports/Sales+Information/Utilities/shipping_invoice_sub_report.rpt
  step_progress_bar()
  WinWait, Report Viewer,,30
  If ErrorLevel
  {
    progress_error(A_LineNumber, "Browser timeout")    
    Goto, End_hotkey
  }
  step_progress_bar()
  Sleep, 3000
  WinActivate
  Send {Tab}%clipboard%
  Send {Tab}{Down}
  Send {Tab 2}{Down 2}{Enter}
  Goto, End_hotkey

  ;-----------------------------------------------------------------------
; [Windows Key + s] Stock Requests from clipboard or highlighted selection
;-------------------------------------------------------------------------
#s::
  create_progress_bar("Stock Requests")
  add_progress_step("Opening web page")
  add_progress_step("Querying part number")
  copy_to_clipboard()

  matches =			; Clear old matches.
  
  ; there's no native function to parse several regex matches, so one has to
  ; reuse the `begin` position parameter to check the full string
  begin = 1
  While begin := RegExMatch(clipboard, "A([A-Z]{3}-[0-9]{5})" ;look for level 1 code of 3 letters and 5 digits
                            , match
                            , begin + StrLen(match))
  {
    part_code%A_Index% := match1
    matches := A_Index
  }
  Loop, %matches%
  {
    add_progress_step("Querying Part Code")
  }
  
  ; open a new window for each match
  Loop, %matches%
  {
    step_progress_bar()
    part := part_code%A_Index%
    Run http://andor.oxinst.com/sageutils/stockrequests/incomplete/edit.asp


  WinWait, Stock Requests,,100
  if ErrorLevel
    {
    progress_error(A_LineNumber)
    Goto, end_hotkey
    }
  WinActivate
  ;while (A_Cursor = "AppStarting")
  Sleep,500
  step_progress_bar()
  SetKeyDelay 10
  Send {tab}{down}{tab}%part%{tab}{Enter}
  }
  
  ; If there are no matches for level 1 codes, assume the selection is
  ; a level 10 code.
  if matches !=
    Goto, End_hotkey  ; level 1 code was in clipboard, so end hotkey.
  
  clipboard:=strip(clipboard)
  Run http://andor.oxinst.com/sageutils/stockrequests/incomplete/edit.asp

  WinWait, Stock Requests,,100
  if ErrorLevel
    {
    progress_error(A_LineNumber)
    Goto, end_hotkey
    }
  WinActivate
  ;while (A_Cursor = "AppStarting")
  Sleep,500
  step_progress_bar()
  SetKeyDelay 10
  Send {tab}{down}{down}{tab}%clipboard%{tab}{tab}{Enter}
  Goto, end_hotkey   
  
  
;-----------------------------------------------------------------------
; [Windows Key + t] Ticket search from clipboard or Outlook e-mail title
;-----------------------------------------------------------------------
#t::
create_progress_bar("Ticket search")
add_progress_step("Extracting Ticket ID")
add_progress_step("Opening ticket in SalesLogix")
add_progress_step("Opening contact in SalesLogix")
copy_to_clipboard()
step_progress_bar()
ticket := get_ticket_number_from_outlook_subject()
if ticket = %NONE_VALUE%
{
  ; Open contact if no ticket number found, so that a new ticket can
  ; be created.
  step_progress_bar()
  step_progress_bar()


  contact_email := get_email_from_clipboard()


  if contact_email = %NONE_VALUE%
    contact_email := get_contact_email_from_outlook_subject()
  
  if contact_email = %NONE_VALUE%
  {
    MsgBox,, Ticket search, No ticket number found in clipboard or e-mail title
    Goto, end_hotkey
  }
  else
  {
    open_contact_by_email(contact_email)
    Goto, end_hotkey
  }
}
step_progress_bar()
open_ticket(ticket)
Goto, end_hotkey

;-----------------------------
; [Windows Key + V] Date paste
;-----------------------------
#v::

;FIXME - Pressing Cancel on dialogue box results in blank name in AHK.ini. It should result in no file being created.
create_progress_bar("Date stamp")
IniRead, INITIALS, %INI_FILE%, Timestamp, initials, NONE_VALUE
if INITIALS = NONE_VALUE
{
  InputBox INITIALS, New Timestamp user, Enter your name or initials - the Date stamp will then appear as "dd-MMM-yyyy YourName" when Win + V is pressed:
  IniWrite %INITIALS%, %INI_FILE%, Timestamp, initials
  if ErrorLevel
  {
    MsgBox Error: Could not write %INITIALS% to %INI_FILE%
    Goto, end_hotkey
  }
}
TimeVar := A_Now
FormatTime, TimeVar, A_Now, dd-MMM-yyyy
SetKeyDelay, 0
Sleep 500
Send %TimeVar% %INITIALS%{Enter 2}
Send Summary:{Enter 3}
Send --------------------------------------------------------------------------------{Enter}
Send Original E-mail:{Enter 3}
Send ********************************************************************************{Enter}
SetKeyDelay, 10			; reset to default value
Goto, end_hotkey





;--------------------------------------------------------------------------
; [Windows Key + z] Bugzilla search from clipboard or highlighted selection
;--------------------------------------------------------------------------
#z::
create_progress_bar("Bugzilla search")
add_progress_step("Reading bug number from selection")
add_progress_step("Opening web link")
copy_to_clipboard()  
step_progress_bar()
found := RegExMatch(clipboard, "(\d+)", bug)
if found
{
  step_progress_bar()
  Run http://be-qa-01/show_bug.cgi?id=%bug1%
}
Goto, end_hotkey

GuiClose:
Gui, Destroy
Return


GuiEscape:
Gui, Destroy
Return


end_hotkey_with_error:
end_hotkey:
kill_progress_bar()
Return