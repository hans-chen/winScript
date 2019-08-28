;Global Variables
global KEY_CTRL_C := 3
global KEY_CTRL_K := 11
global KEY_CTRL_L := 12
global KEY_CTRL_S := 19

;group including windows ignore this ahk
GroupAdd, IGNOREWINDOWS, ahk_class VanDyke Software - SecureCRT	;ignore shortcuts to SecureCRT
GroupAdd, IGNOREWINDOWS, ahk_class Chrome_WidgetWin_1			;ignore shortcuts to Visual Studio Code
GroupAdd, IGNOREWINDOWS, ahk_class XLMAIN						;ignore shortcuts to Excel

;group including 3D design windows need to redifine mouse middle click
GroupAdd, 3DDESIGNWINDOWS, ahk_class HCS16139P
GroupAdd, 3DDESIGNWINDOWS, ahk_class Qt5QWindowIcon

;Quick launch menu definition
Menu, QuickRun, Add, &Everything, QuickRunMenuHandler
Menu, QuickRun, Add, &Keyshot, QuickRunMenuHandler
Menu, QuickRun, Add, C&reo, QuickRunMenuHandler
Menu, QuickRun, Add, Easy&Set, QuickRunMenuHandler
Menu, QuickRun, Add, Speed&Commander, QuickRunMenuHandler
Menu, QuickRun, Add, &Visual Studio Code, QuickRunMenuHandler
Menu, QuickRun, Add, &MindManager, QuickRunMenuHandler
return

QuickRunMenuHandler:
if (A_ThisMenuItem == "&Everything")
	Run, % "C:\Program Files\Everything\Everything.exe", , Max
else if (A_ThisMenuItem == "&Keyshot")
	Run, % "C:\Program Files\KeyShot7\bin\keyshot.exe", % "D:\Docs\CAD\", Max
else if (A_ThisMenuItem == "C&reo")
	Run, % "C:\Program Files\PTC\Creo 5.0.0.0\Parametric\bin\parametric.exe", % "D:\Docs\CAD\", Max
else if (A_ThisMenuItem == "Easy&Set")
	Run, % "C:\Users\haiha\AppData\Local\EasySet\EasySet.exe"
else if (A_ThisMenuItem == "Speed&Commander")
	Run, % "D:\Docs\Tools\SpeedCommander\SpeedCommanderPortable.exe"
else if (A_ThisMenuItem == "&Visual Studio Code")
	Run, % "C:\Program Files\Microsoft VS Code\Code.exe"
else if (A_ThisMenuItem == "&MindManager")
	Run, % "C:\Program Files (x86)\Mindjet\MindManager 18\MindManager.exe"
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

;roaming shortcuts 

#InputLevel 5			;higher input level

;text editing cut/copy/paste
^k::	Send +{End}^x	;Shift+End then Ctrl+x

;text editing insert/add/delete
^o::	Send {Enter}{Up}{End}
^d::	Send {Delete}

;Shortcuts for Outlook
#IfWinActive ahk_class rctrl_renwnd32

TooltipDisplay(tips, timeout:=1000)
{
	ToolTip, % tips
	SetTimer, ClearToolTip, % 0-timeout
	return

ClearToolTip:
	ToolTip
	return
}

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
	TooltipDisplay("L => Reply to all`nK => Reply`nJ => Forward`nCtrl-L => Catogorize Handled`nM => Catogorize Pending")
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
	TooltipDisplay("L => Inbox`nJ => To be vetted`nK => Sent`nM => Follow`nN => Pending")
	key := SecondKeyInput()
	if (key == "l")				 							;Jump to Inbox
		Send ^+{Tab}hp{Enter}
	else if (key == "j")				 					;Jump to To be vetted
		Send ^+{Tab}hg1{Enter}
	else if (key == "k")									;Jump to Sent
		Send ^+{Tab}hv{Enter}
	else if (key == "m")									;Jump to Follow
		Send ^+{Tab}hg2{Enter}
	else if (key == "n")									;Jump to Pending
		Send ^+{Tab}hg3{Enter}

	return
}

SecondShortcuts_CtrlM()
{
	TooltipDisplay("K => Signature 1 side window`nL => Signature 1 popup window`nJ => HTML`nM => Search From`nN => Search To")
	key := SecondKeyInput()
	if (key == "l")				 							;Select first signature in popup editing window
		Send !has{Enter}
	else if (key == "j")				 					;Select HTML on Format tab
		Send !oth
	else if (key == "k")									;Select first signature in side editing window
		Send !e2as{Enter}
	else if (key == "m")									;Jump to Search From
		Send ^e{Tab 3}
	else if (key == "n")									;Jump to Search To
		Send ^e{Tab 5}

	return
}

^l::	SecondShortcuts_CtrlL()								;Messages
^j::	SecondShortcuts_CtrlJ()								;Folders
^m::	SecondShortcuts_CtrlM()								;Editting & Search

;Moving focus
^p::	Send {Up}
^n::	Send {Down}

;end of Outlook

;start of Wechat

#IfWinActive ahk_class WeChatMainWndForPC

^e::	MouseClick, left, 290, 45							;search box
^r::	MouseClick, left, 2470, 1245						;call

;Moving focus
^p::	Send {Up}
^n::	Send {Down}

;end of Wechat

;Mouse middle key for Creo and keyshot
;#IfWinActive, ahk_group 3DDESIGNWINDOWS

;XButton1::	MButton
;XButton2::	MButton
