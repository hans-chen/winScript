;Global Variables
global KEY_CTRL_C := 3
global KEY_CTRL_K := 11
global KEY_CTRL_L := 12
global KEY_CTRL_M := 13
global KEY_CTRL_S := 19

;group including windows ignore this ahk
GroupAdd, IGNOREWINDOWS, ahk_class VanDyke Software - SecureCRT	;ignore shortcuts to SecureCRT
GroupAdd, IGNOREWINDOWS, ahk_exe Code.exe						;ignore shortcuts to Visual Studio Code
GroupAdd, IGNOREWINDOWS, ahk_class XLMAIN						;ignore shortcuts to Excel
GroupAdd, IGNOREWINDOWS, ahk_class Emacs						;ignore shortcuts to Emacs
GroupAdd, IGNOREWINDOWS, ahk_exe pycharm64.exe					;ignore shortcuts to Visual Studio Code

;group including 3D design windows need to redifine mouse middle click
GroupAdd, 3DDESIGNWINDOWS, ahk_class HCS16139P
GroupAdd, 3DDESIGNWINDOWS, ahk_class Qt5QWindowIcon

;Quick launch menu definition
Menu, QuickRun, Add, &Emacs, QuickRunMenuHandler
Menu, QuickRun, Add, Every&thing, QuickRunMenuHandler
Menu, QuickRun, Add, &Visual Studio Code, QuickRunMenuHandler
Menu, QuickRun, Add, E&xact, QuickRunMenuHandler
return

QuickRunMenuHandler:
if (A_ThisMenuItem == "Every&thing")
	Run, % "C:\Program Files\Everything\Everything.exe", , Max
else if (A_ThisMenuItem == "&Emacs")
	Run, % "D:\Programs\Emacs\x86_64\bin\runemacs.exe", % "D:\Docs\org", Max
else if (A_ThisMenuItem == "&Visual Studio Code")
	Run, % "C:\Program Files\Microsoft VS Code\Code.exe"
else if (A_ThisMenuItem == "E&xact")
	Run, % "D:\Docs\Remote Desktop\Exact0.RDP"
return

;Global keys

#m::	Menu, QuickRun, Show

#!r:: Reload					;windows+alt+r to reload script

#!c:: dev_CheckInfo()
dev_CheckInfo()
{
	tooltip
	WinGet, Awinid, ID, A ; cache active window unique id
	WinGetClass, class, ahk_id %Awinid%
	WinGetTitle, title, ahk_id %Awinid%
	WinGetPos, x,y,w,h, ahk_id %Awinid%
	WinGet, pid, PID, ahk_id %Awinid%
	WinGet, exepath, ProcessPath, ahk_id %Awinid%
	ControlGetFocus, focusNN, ahk_id %Awinid%
	ControlGet, focus_hctrl, HWND, , %focusNN%, ahk_id %Awinid%
	
	CoordMode, Mouse, Screen
	MouseGetPos, mxScreen, myScreen
	
	CoordMode, Mouse, Window
	MouseGetPos, mxWindow, myWindow, , classnn
	
	MsgBox, % msgboxoption_IconInfo, ,
	(
		The Active window class is "%class%" (Hwnd=%Awinid%)
		Title is "%title%"
		Position: x=%x% , y=%y% , width=%w% , height=%h%

		Current focused classnn: %focusNN%
		Current focused hctrl: ahk_id=%focus_hctrl%

		Process ID: %pid%
		Process path: %exepath%

		Mouse position: In-window: (%mxWindow%,%myWindow%)  `; In-screen: (%mxScreen%,%myScreen%)

		ClassNN under mouse is "%classnn%"	
	)
}

#IfWinNotActive, ahk_group IGNOREWINDOWS

;emacs shortcuts

#InputLevel 5			;higher input level

;text editing cut/copy/paste
^k::	Send +{End}^x	;Shift+End then Ctrl+x
!d::	Send ^+{Right}^x	;Shift+End then Ctrl+x

;text editing insert/add/delete
^o::	Send {Enter}{Up}{End}
^d::	Send {Delete}

;Moving
^p::	Send {Up}
^n::	Send {Down}
^b::	Send {Left}
^f::	Send {Right}
^a::	Send {Home}
^e::	Send {End}

;Shortcuts for Outlook
#IfWinActive ahk_class rctrl_renwnd32

SecondKeyInput()
{
	Suspend, On
	Input, key, L1 T1 M
	Suspend, Off

	okey := Ord(key)

	if okey > 24		;not ctrl- key
		return key
	else
		return okey
}

SecondShortcuts_CtrlL()
{
	;TooltipDisplay("L => Reply to all`nK => Reply`nJ => Forward`nCtrl-L => Catogorize Handled`nM => Catogorize Pending")
	key := SecondKeyInput()
	if (key == "l")				 							;Reply to all
		Send !hra
	else if (key == "k")				 					;Reply
		Send !hrp
	else if (key == "j")									;Forward
		Send !hfw
	else if (key == KEY_CTRL_L)								;Catogorize Handled
		Send !hy3
	else if (key == "m")									;Catogorize Pending
		Send !hy1

	return
}

SecondShortcuts_CtrlJ()
{
	;TooltipDisplay("L => Inbox`nJ => To be vetted`nK => Sent`nM => Follow`nN => Pending")
	key := SecondKeyInput()
	if (key == "k")				 							;Jump to Inbox
		Send ^+{Tab}hi{Enter}
	else if (key == "j")				 					;Jump to To be vetted
		Send ^+{Tab}hg1{Enter}
	else if (key == "s")									;Jump to Sent
		Send ^+{Tab}hs{Enter}
	else if (key == "l")									;Jump to Follow
		Send ^+{Tab}hg2{Enter}

	return
}

SecondShortcuts_CtrlM()
{
	;TooltipDisplay("K => Signature 1 side window`nL => Signature 1 popup window`nJ => HTML`nM => Search From`nN => Search To")
	key := SecondKeyInput()
	if (key == "l")				 							;Select first signature in popup editing window
		Send !has{Enter}
	else if (key == "j")				 					;Select HTML on Format tab
		Send !oth
	else if (key == "k")									;Select first signature in side editing window
		Send !e2as{Enter}
	else if (key == KEY_CTRL_M)								;Jump to Search
		Send !q1

	return
}

^l::	SecondShortcuts_CtrlL()								;Messages
^j::	SecondShortcuts_CtrlJ()								;Folders
^m::	SecondShortcuts_CtrlM()								;Editting & Search

;end of Outlook

;;;start of Wechat
#IfWinActive ahk_class WeChatMainWndForPC

!q::	MouseClick, left, 200, 45							;search box position, (x,y)
^r::														;call or pick up call
Send +{Tab}{Right 4}{Enter}	

#IfWinActive ahk_class AudioWnd

^h::	Send {Tab 5}{Enter}									;hang up call

#IfWinActive ahk_class VoipTrayWnd

^r::	Send {Tab 3}{Enter}									;pick up call

;;;end of Wechat

;;;start of Wework
#IfWinActive ahk_class WeWorkWindow

!q::	MouseClick, left, 200, 45							;search box position, (x,y)
^r::	MouseClick, left, 725, 1227							;call

;;;end of Wework

;;;start of Teams
#IfWinActive ahk_exe Teams.exe

!q::	Send ^e												;search box

;;;end of Teams

;;;start of Chrome
#IfWinActive ahk_exe chrome.exe

!q::	Send {F6}											;search and address bar
^s::	Send {F3}											;search in page

;;;end of Chrome

;Mouse middle key for Creo and keyshot
;#IfWinActive, ahk_group 3DDESIGNWINDOWS

;XButton1::	MButton
;XButton2::	MButton
