#NoTrayIcon
#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SDCWorkbench.ico
#AutoIt3Wrapper_Outfile=Comcast_IE_Reset.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Comment=Developed by support.com IT
#AutoIt3Wrapper_Res_Description=Internet Explorer Reset tool for the Comcast Tenant.
#AutoIt3Wrapper_Res_Fileversion=1.4.1.1
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Obfuscator=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Version 1.4.1.0
;Written by Scott Stanley

#include "_ProgressGUI.au3"		;This is a UDF include.
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>

Opt("GUIOnEventMode", 1)		;Set the GUI to Event mode.

ConnectivityTest("fast.support.com",0x000000,0x8dc63f)	;Ensure connectivity to our registry file server.

$tempRegFile = EnvGet("TEMP") & "\IES.REG"	;Set up a temporary file for use in resetting IE

if msgbox(36,"Warning!","This application will close all open Internet Explorer windows." & @CRLF & "Do you wish to proceed?") = 7 Then
	Exit
EndIf

While ProcessExists("iexplore.exe") ;Close all Instances of Internet Explorer
      ProcessClose("iexplore.exe")
  Wend

;Back up All Internet Explorer Shortcuts, if it fails notify the user.
If not DirCopy(EnvGet("USERPROFILE") & "\Favorites",EnvGet("USERPROFILE") & "\Fav.bak",1) Then
	$okCancel=MsgBox(33, "Warning!","Unable to Back up your Favorites. Click Ok to Proceed with Reset and Cancel to cancel.")
	If $okCancel = 2 Then
		Exit
	EndIf
EndIf

;Check which version of IE, and attempt to upgrade if IE6 or 7 are found.
Switch GetIEVersion()
	Case "6","7" ;IE 6 or 7 is found
		Switch @OSArch ;Determine 64 or 32 bit OS. Stop the program if IA64.
			Case "X86" ;Determine if 32-bit Win XP or Win Vista and get appropriate upgrade. It is impossible for Windows 7 to have IE6 or IE7 so there is no need to account for it.
				If @OSVersion = "WIN_XP" Then
					MsgBox(64, "Reset Failed", "Unsupported Version - Internet Explorer will attempt to upgrade now. System will Reboot, please save your work.")
					DownloadAndRun("http://fast.support.com/download/Comcast/autoit/IE8/IE8-WindowsXP-x86-ENU.exe", "IE8.exe", "Downloading Internet Explorer 8", "/passive")
					Exit
				ElseIf @OSVersion = "WIN_VISTA" Then
					MsgBox(64, "Reset Failed", "Unsupported Version - Internet Explorer will attempt to upgrade now. System will Reboot, please save your work.")
					DownloadAndRun("http://fast.support.com/download/Comcast/autoit/IE8/IE8-WindowsVista-x86-ENU.exe", "IE8.exe", "Downloading Internet Explorer 8", "/passive")
					Exit
				Else
					MsgBox(48, "Reset Failed", "Unsupported OS Version")
					Exit
				EndIf
			Case "X64" ;Determine if 64-bit Win XP or Win Vista and get appropriate upgrade. It is impossible for Windows 7 to have IE6 or IE7 so there is no need to account for it.
				If @OSVersion = "WIN_XP" OR @OSVersion = "WIN_2003" Then
					MsgBox(64, "Reset Failed", "Unsupported Version - Internet Explorer will attempt to upgrade now. System will Reboot, please save your work.")
					DownloadAndRun("http://fast.support.com/download/Comcast/autoit/IE8/IE8-WindowsServer2003-x64-ENU.exe", "IE8.exe", "Downloading Internet Explorer 8", "/passive")
					Exit
				ElseIf @OSVersion = "WIN_VISTA" Then
					MsgBox(64, "Reset Failed", "Unsupported Version - Internet Explorer will attempt to upgrade now. System will Reboot, please save your work.")
					DownloadAndRun("http://fast.support.com/download/Comcast/autoit/IE8/IE8-WindowsVista-x64-ENU.exe", "IE8.exe", "Downloading Internet Explorer 8", "/passive")
					Exit
				Else
					MsgBox(48, "Reset Failed", "Unsupported OS Version")
					Exit
				EndIf
			Case Else ;We don't support IA64 or other Architectures.
				MsgBox(48, "Reset Failed", "Unsupported Version of Internet Explorer or Operating System. IE Version: " & GetIEVersionFull() & ". Please contact IT or file a FAST ticket")
				ShellExecute("iexplore", "-new https://fast.support.com")
				Exit
		EndSwitch
	;All versions of 8 and 9 not including RCs and Betas
	Case "8","9","10"
		;Open inetcpl.cpl to the advanced tab
		Run("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,Internet,6")
		;Reset it
		WinWaitActive("Internet Properties", "Advanced")
		if GetIEVersionFull() = "8.0.6001.19088" Then ;There is 1 version of IE that we've found where sending the reset key sequence doesn't properly work. This accounts for it.
			MsgBox(48,"Warning! User Input Required!","You will need to manually click the reset button on the Internet Explorer Properties window that is open on your desktop after clicking Ok on this window to proceed.")
		EndIf
		Send("!s") ; this breaks on 8.0.6001.19088 - See above If statement for further details.
		WinWaitActive("Reset Internet Explorer Settings")
		send("!p")
		send("!r")
		;this loop waits for the IE reset to complete.
		WinWait('Reset Internet Explorer Settings')
			While Not ControlCommand('Reset Internet Explorer Settings','', 'Button1','IsEnabled','')
				Sleep(500)
			WEnd
		Send('!c')
		;Re-Check IE Version (8/9) and download appropriate settings file.
		If GetIEVersion() = "8" Then
			InetGet ("https://fast.support.com/download/Comcast/autoit/Comcast_IE8_Settings.reg", $tempRegFile,1,1)
			$proBar = _ProgressGUI("Configuring Internet Explorer",0,14,"Arial",290,100,"0x000000","0x8dc63f")
			While InetGetInfo() > 0
				Sleep(100)
			WEnd
			Run("Regedit /s " & $tempRegFile)
			GUIDelete($proBar[0])
		Else
			InetGet ("https://fast.support.com/download/Comcast/autoit/Comcast_IE9_Settings.reg", $tempRegFile,1,1)
			$proBar = _ProgressGUI("Configuring Internet Explorer",0,14,"Arial",290,100,"0x000000","0x8dc63f")
			While InetGetInfo() > 0
				Sleep(100)
			WEnd
			Run("Regedit /s " & $tempRegFile)
			GUIDelete($proBar[0])
		EndIf
	;Fallback Clause. If the IE version is unrecognized, errors out.
	Case Else
		MsgBox(48, "Reset Failed", "Unsupported Version of Internet Explorer or Operating System. IE Version: " & GetIEVersionFull() & ". Please contact IT or file a FAST ticket")
		ShellExecute("iexplore", "-new https://fast.support.com")
		Exit
EndSwitch
;Restore 2 registry keys that get mangled when resetting IE on some versions of windows
if @OSArch = 'X64' Then
	RegWrite('HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0\win32',"","REG_SZ",'C:\Windows\SysWOW64\ieframe.dll')
	RegWrite('HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0\win64',"","REG_SZ",'C:\Windows\System32\ieframe.dll')
Else
	RegWrite('HKEY_CLASSES_ROOT\TypeLib\{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}\1.1\0\win32',"","REG_SZ",'C:\Windows\System32\ieframe.dll')
EndIf

;Restore Internet Explorer Shortcuts, if it fails notify the user.
If not DirCopy(EnvGet("USERPROFILE") & "\Fav.bak",EnvGet("USERPROFILE") & "\Favorites",1) Then
	MsgBox(48, "Warning!","Unable to Restore your Favorites. Opening what should be the backup folder now.")
	Run("explorer.exe " & EnvGet("USERPROFILE") & "\Fav.bak");
EndIf
MsgBox(0, "Reset Complete", "Internet Explorer Reset is complete.")
FileDelete($tempRegFile)

;Functions that grab the IE version directly from the iexplore.exe file.
Func GetIEVersion()
	;Local $ieVersion = StringLeft(FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe"), 1)
	Local $ieVersion = StringSplit(FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe"), ".")
	return $ieVersion[1]
EndFunc

Func GetIEVersionFull()
	Local $ieVersionFull = FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe")
	return $ieVersionFull
EndFunc

Func DownloadAndRun($remotefile, $localfilename, $GUIMsg, $args)
Local $tempFile = EnvGet("TEMP") & $localfilename
	InetGet($remotefile, $tempFile, 1,1)
		Local $dlBar = _ProgressGUI($GUIMsg,0,14,"Arial",310,100,"0x000000","0x8dc63f")
			While InetGetInfo() > 0
				Sleep(100)
			WEnd
		GUIDelete($dlBar[0])
		ShellExecute($tempFile, $args)
EndFunc

;Connectivity Test Function - May want to consider adding more options to this function and moving to an include, as it's pretty long.
Func ConnectivityTest($site, $bgcolor, $fgcolor)
Local $failureCount = 0

;Set up a small GUI window for the status of the connectivity test.
$iGUI=GUICreate("Internet Connectivity Test",300,80)
GUISetOnEvent($GUI_EVENT_CLOSE, "inetCancel")

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

Func inetCancel()
	Exit
EndFunc