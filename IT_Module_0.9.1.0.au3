#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SDCWorkbench.ico
#AutoIt3Wrapper_Outfile=Workbench_IT.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Fileversion=0.9.1.0
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Obfuscator=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#include <GUIConstantsEx.au3>
#include "_ProgressGUI.au3"
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include "SDCWorkbench.au3"

Opt("GUIOnEventMode", 1)
Opt("TrayIconHide",1)	;Hide the System Tray icon for this program.

if $CmdLine[0] > 0 Then
	switch $CmdLine[1]
		case 'clip'
			ClipPut('https://fast.support.com/download/SDCWorkbench/Workbench_Launcher.exe')
		case 'ucn'
			$CheckProc = 0
				While $CheckProc = 0
					$Kill = "ok"
					if (ProcessExists("myagent.exe") or ProcessExists("Ninjato.exe") or ProcessExists("EyeBeam.exe")) Then
						if msgbox(5,"Please exit the following applications", "1. eyeBeam Softphone (eyeBeam.exe)" & @LF & "2. InContact Agent (MyAgent.exe)" & @LF & "3. Ninjato.exe") = 2 Then
							ExitLoop
						EndIf
					Else
						$CheckProc = 1
					EndIf
				WEnd
			if $CheckProc = 1 Then
				RunFromReg("HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Support.com Ninjato|UninstallString",1)
				if @OSVersion = "WIN_XP" Then
					DirRemove(EnvGet("USERPROFILE") & "\Application Data\supportdotcom",1)
					DirRemove(EnvGet("USERPROFILE") & "\Local Settings\Application Data\Support.com,_Inc",1)
					RegDelete("HKCU\Software\MyACD.com")
				Else
					DirRemove(EnvGet("APPDATA") & "\supportdotcom",1)
					DirRemove(EnvGet("USERPROFILE") & "\AppData\Local\support.com,_Inc",1)
					RegDelete("HKCU\Software\MyACD.com")
				EndIf
			EndIf
		case 'vpn'
			$CheckProc = 0
				While $CheckProc = 0
					$Kill = "ok"
					if (ProcessExists("cscan.exe") or ProcessExists("vpnui.exe")) Then
						ProcessClose("cscan.exe")
						if msgbox(5,"Please exit the following applications", "1. Cisco AnyConnect Client (vpnui.exe)" & @LF & "2. Cisco Scanner (cscan.exe)") = 2 Then
							ExitLoop
						EndIf
					Else
						$CheckProc = 1
					EndIf
				WEnd
			if $CheckProc = 1 Then
				RunFromReg("HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\{668842FC-6827-4B6F-82BF-3828BE6D3007}|UninstallString",1)
				if @OSVersion = "WIN_XP" Then
					DirRemove(EnvGet("USERPROFILE") & "\Application Data\Cisco",1)
					DirRemove(EnvGet("USERPROFILE") & "\Local Settings\Application Data\Cisco",1)
					RegDelete("HKCU\Software\Cisco")
					RegDelete("HKCU\Software\Cisco Systems")
					RegDelete("HKLM\Software\Cisco")
				Else
					DirRemove(EnvGet("APPDATA") & "\Cisco",1)
					DirRemove(EnvGet("USERPROFILE") & "\AppData\Local\Cisco",1)
					RegDelete("HKCU\Software\Cisco")
					RegDelete("HKCU\Software\Cisco Systems")
					RegDelete("HKLM\Software\Cisco")
				EndIf
			EndIf

		case 'info'
			SupportInfo()
		case 'ssave'
			CheckScreenSaver()
	EndSwitch
EndIf

Func SupportInfo()						;Troubleshooting Info
	Local $failureCount = 0
	Local $ping
	dim $supportGUI

	$iGUI = GUICreate("IT Support Information", 300, 80)	;Sets up a GUI with a message telling the user we're collecting some data.
	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
	GUISetBkColor(0x000000)
	$statusmessage = GUICtrlCreateLabel("Gathering Data...", 1, 30, 298)
	GUICtrlSetStyle($statusmessage, $SS_Center)
	GUICtrlSetColor($statusmessage, $textColor)
	GUISetState(@SW_SHOW)

	;Open our debug file in write mode and add some basic computer info and a header.
	$debugfile = FileOpen(@ScriptDir & "\debug_log.txt", 2)
	FileWrite($debugfile, "=============BEGIN SUPPORT.COM IT DEBUG INFO=============" & @CRLF)
	FileWrite($debugfile, "" & @ComputerName & " - " & @OSVersion & " " & @OSArch & " " & @OSServicePack & " -- Current Time: " & _Now() & @CRLF)

	;NAC Test
	If Not ProcessExists("NacAgent.exe") Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: NAC Assessment Agent Not Running" & @CRLF)
	Else
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: NAC Assessment Agent Present and Running" & @CRLF)
	EndIf

	;VPN Connectivity Test
	$ping = Ping("10.51.1.42", 900)
	While @error ;if unable to contact the site selected, retry.
		If $failureCount < 5 Then ;up to 5 times
			Sleep(1000) ;at 1s intervals with a timeout of 900ms
			$failureCount = $failureCount + 1 ;increase the # of failures
		Else
			FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: Unable to Confirm VPN Connectivity at 10.51.1.42" & @CRLF)
			ExitLoop
		EndIf
		$ping = Ping("10.51.1.42", 900) ;This must be dead last in the loop otherwise @error will change erroneously.
	WEnd
	If $failureCount < 5 Then ;If the last loop exitted with less than 5 failures we successfully connected
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: VPN Connectivity to 10.51.1.42 Confirmed" & @CRLF)
	EndIf

	;NSLookup on www-ninjato.support.com
	TCPStartup()
	FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: www-ninjato.support.com resolves to the following IP Address: [" & TCPNameToIP("www-ninjato.support.com") & "]" & @CRLF)
	TCPShutdown()

	;Check for full version of .NET 4
	$dotnetinstalled = RegRead('HKEY_LOCAL_MACHINE\Software\Microsoft\NET Framework Setup\NDP\v4\Full', 'Install')
	If @error <> 0 Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: Unable to access .NET v4 (full) Framework Registry Key" & @CRLF)
	ElseIf $dotnetinstalled == '0' Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: .NET v4 Framework(full) Not Installed" & @CRLF)
	ElseIf $dotnetinstalled == '1' Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: .NET v4 Framework(full) Installed" & @CRLF)
	EndIf

	;Check for client only version of .NET 4
	$dotnetinstalled = RegRead('HKEY_LOCAL_MACHINE\Software\Microsoft\NET Framework Setup\NDP\v4\Client', 'Install')
	If @error <> 0 Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: Unable to access .NET v4 (client) Framework Registry Key" & @CRLF)
	ElseIf $dotnetinstalled == '0' Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: .NET v4 Framework(client) Not Installed" & @CRLF)
	ElseIf $dotnetinstalled == '1' Then
		FileWrite($debugfile, "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]: .NET v4 Framework(client) Installed" & @CRLF)
	EndIf

	;Close the debug file
	FileClose($debugfile)
	;Close our Gathering Info window
	GUIDelete($iGUI)
	;Open up a textbox display GUI element and set up the event for closing it.
	$filedisplay = FileRead(@ScriptDir & "\debug_log.txt")
	$supportGUI = GUICreate("Support Information", 640, 480)
	$supportbox = GUICtrlCreateEdit($filedisplay, 0, 0, 640, 480, $ES_AUTOVSCROLL + $WS_VSCROLL + $ES_READONLY)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked", $supportGUI)
	GUISetState(@SW_SHOW)
	while 1
		WEnd
EndFunc   ;==>SupportInfo

Func CLOSEClicked()
	Exit
EndFunc   ;==>CLOSEClicked