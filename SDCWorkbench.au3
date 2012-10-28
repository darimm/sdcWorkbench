#include-once

Const $textColor = 0x8dc63f

Func getHTTPfilename($tempString) ;returns the filename of an http(s) link stripped of illegal characters
	Return StringReplace(StringReplace(StringReplace(StringTrimLeft($tempString, StringInStr($tempString, "/", Default, -1)), "?", ""), "%", ""), "*", "")
EndFunc   ;==>getHTTPfilename

Func CheckScreenSaver()				;this should probably get moved out to the SDCWorkbench.au3 include.
	RegRead('HKEY_CURRENT_USER\Control Panel\Desktop', 'SCRNSAVE.EXE') ;If this doesn't exist it means a screensaver has never been set.
	If @error <> 0 Then
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaverIsSecure', "REG_SZ", '1') ;password protected
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaveTimeOut', "REG_SZ", '1140') ;19 minutes
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaveActive', "REG_SZ", '1') ;Active
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'SCRNSAVE.EXE', "REG_SZ", @SystemDir & '\scrnsave.scr') ;Blank Screen Screensaver
	Else ;Screensaver is set but turned off, turn it back on.
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaverIsSecure', "REG_SZ", '1') ;password protected
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaveTimeOut', "REG_SZ", '1140') ;19 minutes
		RegWrite('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaveActive', "REG_SZ", '1') ;Active
	EndIf
EndFunc   ;==>CheckScreenSaver

Func GetIEVersionFull()	;This should definitely get moved out to the SDCWorkbench.au3 include.
	Local $ieVersionFull = FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe")
	Return $ieVersionFull
EndFunc   ;==>GetIEVersionFull

Func _UserInput($input, $string) ;This function strips out the $USERINPUT: and replaces it with the actual user's input
	return StringLeft($string,StringinStr($string,"$USERINPUT:")-1) & $input & StringRight($string,Stringlen($string) - StringinStr($string,"$",0,2))
EndFunc  ;==>_UserInput

;Rewritten to preserve filenames or approximations, depends on getHTTPfilename function
Func DownloadAndRun($remotefile, $GUIMsg, $args = "")
	Local $tempFile = EnvGet("TEMP") & "\" & getHTTPfilename($remotefile)
	InetGet($remotefile, $tempFile, 1, 1)
	Local $dlBar = _ProgressGUI($GUIMsg, 0, 14, "Arial", 290, 100, "0x000000", $textColor)
	While InetGetInfo() > 0
		Sleep(100)
	WEnd
	GUIDelete($dlBar[0])
	ShellExecute($tempFile, $args)
EndFunc   ;==>DownloadAndRun

Func RunFromReg($reg_string, $wait = 0)							;Run something from a registry key, optionally wait for completion
	$tempParse=StringSplit($reg_string,"|")						;Using |s to delimit options for this from the configfile
	if $tempParse[0] = '2' then									;If there are 2 options use this loop
		$button_do = RegRead($tempParse[1],$tempParse[2])
		if @error = 0 Then
			if $wait = 0 Then
				Run($button_do)
			Else
				RunWait($button_do)
			EndIf
		EndIf
	ElseIf $tempParse[0] = '3' Then								;If there are 3 options use this loop
		$button_do = RegRead($tempParse[1],$tempParse[2])
		if @error = 0 Then
			if StringRight($button_do, 1) <> '\' Then
				if $wait = 0 Then
					Run($button_do & '\' & $tempParse[3])
				Else
					RunWait($button_do & '\' & $tempParse[3])
				EndIf
			Else
				if $wait = 0 Then
					Run($button_do & $tempParse[3])
				Else
					RunWait($button_do & $tempParse[3])
				EndIf
			EndIf
		EndIf
	EndIf
EndFunc	;==>RunFromReg