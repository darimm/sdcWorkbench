#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SDCWorkbench.ico
#AutoIt3Wrapper_Outfile=Workbench_Login.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Fileversion=0.9.1.0
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Obfuscator=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Version 0.1
;Written by Scott Stanley

#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include "_ProgressGUI.au3"
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <ComboConstants.au3>
#include "SDCWorkbench.au3"

Opt("GUIOnEventMode", 1)

Global $hGUI

If ConnectivityTest("fast.support.com", 0x000000, 0x8dc63f, 4, 0) = 1 Then
	MsgBox(0,"Internet Connectivity","Internet Connection Lost. Login Module will close.")
	Exit
EndIf

;Global $loginConfig = LoadConfig("https://fast.support.com/download/SDCWorkbench/dropdown.cfg")
Global $loginConfig = LoadConfig("http://dl.dropbox.com/u/16751601/New%20folder/dropdown.cfg")
Global $tenantbox
global $changingtheMD5 = 123456

	$hGUI = GUICreate("Support.com WorkBench Login",300,100)
	GUISetBkColor(0x000000) ;Set Background color
	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")				;Set Close event
	$tempArray=StringSplit($loginConfig[1],",")
	$tenantbox = GUICtrlCreateCombo($tempArray[1],30,10,240,-1,$CBS_DROPDOWNLIST)
	For $i = 2 to $loginConfig[0]
		$tempArray=StringSplit($loginConfig[$i],",")
		if $tempArray[0] = 2 Then
			GUICtrlSetData($tenantbox, $tempArray[1])
		EndIf
	Next
	$trainingBox = GUICtrlCreateCheckbox(" ",280,75)
	GUIctrlSetColor(GUICtrlCreateLabel("In Training",220,78),0x8dc63f)

	$myButton = GUICtrlCreateButton("Login",110,60,80)
	GUICtrlSetOnEvent($myButton,"DoButtonStuff")
	$textLabel = GUICtrlCreateLabel("Please Select your Team",1,40,298,0,$SS_CENTER)
	GUICtrlSetColor($textlabel,0x8dc63f)
	GUICtrlSetState($textlabel, $GUI_FOCUS)
	GUISetState(@SW_SHOW)

While 1
WEnd


Func DoButtonStuff()
	For $i = 1 to $loginConfig[0]
		if StringLeft($loginConfig[$i], StringInStr($loginConfig[$i], ",") - 1) = GUICtrlRead($tenantbox) Then
			$tempArray=StringSplit($loginConfig[$i],",")
			if $tempArray[0] = 2 Then
				if _IsChecked($trainingBox) Then
					$tempArray[2] = StringReplace($tempArray[2],"it_tools","it_tools training")
;					DownloadAndRun("http://dl.dropbox.com/u/16751601/Workbench_Enduser.exe","Logging you in...",$tempArray[2])
					DownloadAndRun("https://fast.support.com/download/SDCWorkbench/Workbench_Enduser.exe","Logging you in...",$tempArray[2])
					Exit
				Else
;					DownloadAndRun("http://dl.dropbox.com/u/16751601/Workbench_Enduser.exe","Logging you in...",$tempArray[2])
					DownloadAndRun("https://fast.support.com/download/SDCWorkbench/Workbench_Enduser.exe","Logging you in...",$tempArray[2])
					Exit
				EndIf
			Else
				Msgbox(0,"ERROR","An Unexpected Error has occurred. Please contact the IT department.")
				Exit
			EndIf
		EndIf
	Next
EndFunc

;Connectivity Test Function
Func ConnectivityTest($site, $bgcolor, $fgcolor, $retries = 500, $showwindow = 1)
	Local $failureCount = 0
	Local $ping

	If $showwindow = 1 Then;Set up a small GUI window for the status of the connectivity test.
		$iGUI = GUICreate("Internet Connectivity Test", 300, 80)
		GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
		GUISetBkColor($bgcolor)
		$statusmessage = GUICtrlCreateLabel("Attempting to Authorize...", 1, 30, 298)
		GUICtrlSetStyle($statusmessage, $SS_Center)
		GUICtrlSetColor($statusmessage, $fgcolor)
		GUISetState(@SW_SHOW)
	EndIf
	$ping = Ping($site, 900)

	While @error ; if we're unable to contact the site we've selected, retry, up to 500 times at 1s intervals, with a timeout of 900ms
		If $showwindow = 1 Then
			GUICtrlSetData($statusmessage, "Authorization Failed. Retrying.")
		EndIf
		If $failureCount < $retries Then
			Sleep(1000)
			$failureCount = $failureCount + 1
		Else ; End the program after 500 (or user defined) attempts. This amount is so high to give people time to log into the vpn.
			If $showwindow = 1 Then
				GUICtrlSetData($statusmessage, "Authorization Timed Out. Exiting.")
				Sleep(5000)
			EndIf
		EndIf
		$ping = Ping($site, 900) ; This must be dead last in the loop otherwise @error will change erroneously.
	WEnd
	If $showwindow = 1 Then
		GUICtrlSetData($statusmessage, "Authorization Succeeded! Proceeding.") ; This only happens when we successfully connect
		Sleep(1000)
		GUIDelete($iGUI) ;Remove the status window
	EndIf

	If $failureCount < $retries Then ;If ping test succeeded return 0
		Return 0
	Else
		Return 1 ;otherwise return 1
	EndIf
EndFunc   ;==>ConnectivityTest

Func LoadConfig($configurationFile)
	Return StringSplit(BinaryToString(InetRead($configurationFile, 17)), @CRLF, 1) ; Read Config file into an array of strings
EndFunc   ;==>LoadConfig

Func _IsChecked($iControlID)
    Return BitAND(GUICtrlRead($iControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

;Function to kill the whole program
Func CLOSEClicked()
	Exit
EndFunc   ;==>CLOSEClicked