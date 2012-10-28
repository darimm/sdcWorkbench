#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=Workbench_Launcher.exe
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Fileversion=0.1.0.0
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_icon=SDCWorkbench.ico
#AutoIt3Wrapper_Run_Obfuscator=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
;Version 0.1
;Written by Scott Stanley

#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include "_ProgressGUI.au3"

ConnectivityTest("fast.support.com",0x000000,0x00ff00)

if $CmdLine[0] > 0 Then
	if $CmdLine[1] = '-remove' Then
		DownloadAndRun("https://fast.support.com/download/SDCWorkbench/SDCWB_Uninstall.exe","SDCWB_Uninstall.exe","Removing...")
		Exit
	EndIf
EndIf

if @ScriptDir <> EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench" Then
	DirRemove(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench",1)
	DirCreate(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench")
	InetGet("https://fast.support.com/download/SDCWorkbench/Workbench_Launcher.exe", EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench\Workbench_Launcher.exe", 1,1)
	$dlBar = _ProgressGUI("Installing SDC WorkBench",0,14,"Arial",260,100,"0x000000","0x00ff00")
		While InetGetInfo() > 0
			Sleep(100)
		WEnd
	GUIDelete($dlBar[0])
	FileCreateShortcut(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench\Workbench_Launcher.exe",@DesktopDir & "\Support.com WorkBench.lnk",EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench","","SDC Daily WorkBench Tool")
	DirCreate(@StartMenuCommonDir & "\Programs\supportdotcom\SDCWorkbench")
	FileCreateShortcut(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench\Workbench_Launcher.exe",@StartMenuCommonDir & "\Programs\supportdotcom\SDCWorkbench\Support.com WorkBench.lnk",EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench","","SDC Daily WorkBench Tool")
	FileCreateShortcut(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench\Workbench_Launcher.exe",@StartMenuCommonDir & "\Programs\supportdotcom\SDCWorkbench\Remove Support.com WorkBench.lnk",EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench","-remove","Remove SDC Daily WorkBench Tool")
	Run(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench\Workbench_Launcher.exe")
	Exit
EndIf

DownloadAndRun("https://fast.support.com/download/SDCWorkbench/Workbench_Login.exe","Workbench_Login.exe","Loading...")

;Connectivity Test Function - May want to consider adding more options to this function and moving to an include, as it's pretty long.
Func ConnectivityTest($site, $bgcolor, $fgcolor)
Local $failureCount = 0
Local $ping

;Set up a small GUI window for the status of the connectivity test.
$iGUI=GUICreate("Internet Connectivity Test",300,80)
GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
GUISetBkColor($bgcolor)
$statusmessage=GUICtrlCreateLabel("Attempting to Authorize...",1,30,298)
GUICtrlSetStyle($statusmessage, $SS_Center)
GUICtrlSetColor($statusmessage, $fgcolor)
GUISetState(@SW_SHOW)

$ping = Ping($site,900)

While @error ; if we're unable to contact the site we've selected, retry, up to 500 times at 1s intervals, with a timeout of 900ms
GuiCtrlSetData($statusmessage, "Authorization Failed. Retrying.")
If $failureCount < 500 Then
	Sleep(1000)
	$failureCount = $failureCount+1
Else ; End the program after 500 attempts. This amount is so high to give people time to log into the vpn.
	GuiCtrlSetData($statusmessage, "Authorization Timed Out. Exiting.")
	Sleep(5000)
	Exit
EndIf
$ping = Ping($site,900) ; This must be dead last in the loop otherwise @error will change erroneously.
WEnd

GuiCtrlSetData($statusmessage, "Authorization Succeeded! Proceeding.") ; This only happens when we successfully connect
Sleep(1000)
GUIDelete($iGUI) ;Remove the status window
EndFunc

;This function is pretty self explanatory.
Func DownloadAndRun($remotefile, $localfilename, $GUIMsg, $args = "")
Local $tempFile = EnvGet("TEMP") & "\" & $localfilename
	InetGet($remotefile, $tempFile, 1,1)
		Local $dlBar = _ProgressGUI($GUIMsg,0,14,"Arial",120,100,"0x000000","0x00ff00")
			While InetGetInfo() > 0
				Sleep(100)
			WEnd
		GUIDelete($dlBar[0])
		ShellExecute($tempFile, $args)
EndFunc

;Function to kill the whole program
Func CLOSEClicked()
	Exit
EndFunc
