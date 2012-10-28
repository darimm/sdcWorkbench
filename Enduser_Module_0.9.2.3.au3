#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SDCWorkbench.ico
#AutoIt3Wrapper_Outfile=Workbench_Enduser.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Fileversion=0.9.2.3
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Version 0.9
;Written by Scott Stanley

#include <GUIConstantsEx.au3>
#include "_ProgressGUI.au3"
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include "SDCWorkbench.au3"
#include <String.au3>
#include <WinAPI.au3>

Opt("GUIOnEventMode", 1)
Opt("TrayIconHide",1)	;Hide the System Tray icon for this program.

;Const $baseURL = 'https://fast.support.com/download/SDCWorkbench/' 	;Production
Const $refreshTimer = 180												;Production - 15 minutes on configfile refresh
Const $baseURL = 'http://dl.dropbox.com/u/16751601/New Folder/'			;Dev
;Const $refreshTimer = 2													;Dev - 10 seconds on configfile refresh

;Test Connection, Fetch Logo
TCPStartup()
	if TCPNameToIP("fast.support.com") <> '12.157.178.159' Then
		TCPShutdown()
		msgbox(0,"ERROR","DNS ERROR. Click Ok to Quit.")
		Exit
	Else
		TCPShutdown()
	EndIf

If ConnectivityTest("fast.support.com", 0x000000, $textColor) = 1 Then
	Exit
EndIf

InetGet($baseURL & "logo.gif", @ScriptDir & "\logo.gif", 1, 0)

Global $hGUI, $picGUI, $supportGUI ; SDC Graphics variables
Global $intervals = 0				;This measures passage of time in 5s intervals. Once it reaches $refreshtimer we trigger a configfile update check.
Global $ss_reset = 0				;Toggle 0/1 - Has screensaver been reset by the program?
Global $abs_NumberOfButtons = 0		;This variable tracks the total number of buttons defined in the configuration file, is used to dim an array that contains button actions
Global $NumberOfTabs = 0			;Total #of tabs, again, used for an array for tab controls
Global $TopButtonsPerTab = 0		;Largest # of buttons on any single tab, this determines our starting program height. Width is static
Global $guiVersion = 0				;Configfile version - set to 0 so that the version update mechanics automatically run when we read a config file for the first time
Global $guiHeight = 0				;Self explanitory, initialized here
Global $guiWidth = 480				;Legacy definition, this # can technically be anything, it's  rewritten the first time we read a configfile.
Global $guiX = -1					;Used for redrawing the GUI after a resize
Global $guiY = -1					;Used for redrawing the GUI after a resize

;Load Config
Global $moduleConfig = LoadConfig($baseURL & "buttons.cfg") ;Read Config file into an array of strings
$guiVersion = PrepareConfig($moduleConfig, $NumberOfTabs, $TopButtonsPerTab, $abs_NumberOfButtons) ;Prepare these 3 variables for use in creating the UI
Global $ButtonArray[$abs_NumberOfButtons + 1] ;Buttons
Global $TabArray[$NumberOfTabs + 1] ;Tabs
Global $ButtonRefArray[($abs_NumberOfbuttons*2)+($NumberOfTabs*2)+1];This Array is special, it has to have enough elements to accomodate any button's control ID as an index for the Array
																;Therefore we account for ALL control IDs (tabs and buttons when defining it. This has been stress tested up to 500
																;GUI elements (300+ more than we currently use if all tabs were loaded) so I assert the math for this is correct.
DrawMainGUI()	;Fairly self-explanatory, calls the MainGUI
DrawLogo()		;..and then the logo

While 1									;Main Program Loop
	if $intervals = $refreshTimer Then 	;Every 15 minutes or so check to see if the configfile has been updated (5s[5000ms] * 180 = 15 minutes)
		$moduleConfig = LoadConfig($baseURL & "buttons.cfg") ;Read Config file into an array of strings
		$tempversion = PrepareConfig($moduleConfig, $NumberOfTabs, $TopButtonsPerTab, $abs_NumberOfButtons) ;reload the configfile and check version
		if $guiVersion < $tempversion Then 				;if there's a new version on the server...
			$guiVersion = $tempversion					;set the version running to the version on the server
			GUIDelete($picGUI)							;nuke the old UI
			GUIDelete($hGUI)							;nuke the old UI
			ReDim $ButtonArray[$abs_NumberOfButtons +1]	;Resize the arrays
			ReDim $TabArray[$NumberOfTabs+1]
			ReDim $ButtonRefArray[($abs_NumberOfbuttons*2)+($NumberOfTabs*2)+1] ; See above
			DrawMainGUI()								;Redraw the UI and Logo
			DrawLogo()
		EndIf
		$intervals = 0									;Reset the 15 minute timer
	EndIf
	Sleep(5000)
	If ProcessExists("ssrangsv.exe") And $ss_reset == 0 Then ; check for RANG, and see if screensaver has been reset yet.
		If RegRead('HKEY_CURRENT_USER\Control Panel\Desktop', 'ScreenSaveTimeout') <> '1140' Then ;check for turned off ss
			CheckScreenSaver() ;Turn the screensaver back on.
			$ss_reset = 1 ;Don't do this again.
		EndIf
	EndIf
	If ProcessExists("ssrangsv.exe") = 0 And $ss_reset == 1 Then ;Check to see if rang has been terminated.
		$ss_reset = 0 ; reset screensaver change ability in the event of a future RANG connection
	EndIf
	$intervals = $intervals + 1		;internal timer for GUI update check. +1 every 5s
WEnd								;==>Main Program Loop

;The Following Function corresponds to buttons on the GUI
Func DoButtonStuff()
	$buttonAction = StringSplit($ButtonRefArray[@GUI_CtrlId], ",")		;Access the reference array at the index of the last ControlID used(button pressed) to get the correct action.
	switch $buttonAction[2]												;Element 2 contains the numeric value that determines which action is taken
		case '0'														;This is the 'Shellexec' case - We run it with the default system handler.
			$buttonsplit=Stringsplit($buttonAction[4],"|")
			if $buttonsplit[0] > 1 Then									;Check for commandline arguements in the configfile. Discarding anything that comes after a second pipe, if there is one(malformed)
				if StringInStr($buttonAction[4],"$USERINPUT:") = 0 Then	;Check for $USERINPUT:<Instructions>$
					ShellExecute($buttonsplit[1],$buttonsplit[2])		;No $USERINPUT
				Else
					$requestArray = _stringBetween($buttonAction[4],"$USERINPUT:","$")	;Grab the Instructions and throw them in an array
					$userInput = InputBox("Input Required","This Function Requires additional Information" & @CRLF & @CRLF & $requestArray[0]) ;Ask the user for their input
					if @error = 0 Then													;If the user clicks cancel or gives no input then do nothing and go back to the main loop
						$tempbuttonaction = _UserInput($userinput, $buttonAction[4])	;_UserInput() strips the $USERINPUT:<Blah>$ and replaces it with what the user typed. Defined in SDCWorkbench.au3
						$buttonsplit=Stringsplit($tempbuttonaction,"|")	;We already know there are commandline arguements from earlier, redefine $buttonsplit after the user input replacement
						ShellExecute($buttonsplit[1],$buttonsplit[2])	;And execute.
					EndIf
				EndIf
			Else														;If no commandline arguements are present..
				if StringInStr($buttonAction[4],"$USERINPUT:") = 0 Then	;Check for $USERINPUT
					ShellExecute($buttonAction[4])						;None here. Just run it normally
				Else
					$requestArray = _stringBetween($buttonAction[4],"$USERINPUT:","$")	;Otherwise, same as above, grab the Question to ask the user
					$userInput = InputBox("Input Required","This Function Requires additional Information" & @CRLF & @CRLF & $requestArray[0]) ;Ask it, grab input
					if @error = 0 Then													;User clicks cancel or provides no info go back to the main loop
						$tempbuttonaction = $tempbuttonaction = _UserInput($userinput, $buttonAction[4])	;swap $USERINPUT:<Blah>$ with the actual user input
						ShellExecute($tempbuttonaction)														;And Execute
					EndIf
				EndIf
			EndIf
		case '1'											;This is the Internet Explorer case - shellexec('iexplore', '-new ' + whatever's in the config. User input is supported but not cmdline args
			if StringInStr($buttonAction[4],"$USERINPUT:") = 0 Then		;As above, check for presence of user input
				ShellExecute("iexplore", "-new " & $buttonAction[4])	;If there is none, just open the webpage in IE
			Else
				$requestArray = _stringBetween($buttonAction[4],"$USERINPUT:","$")	;Otherwise, grab the question to ask the user
				$userInput = InputBox("Input Required","This Function Requires additional Information" & @CRLF & @CRLF & $requestArray[0])	;and Ask it
				if @error = 0 Then													;User gives us valid input, if cancel/empty input kicks back to the main program
					ShellExecute("iexplore", "-new " & _UserInput($userinput, $buttonAction[4]))	;Open the webpage, replacing $USERINPUT:<Blah>$ with the actual user's input.
				EndIf
			EndIf
		case '2'											;This is the Download from the internet and run case - used to launch the IT module, and various removal/install tools.
			$buttonsplit=Stringsplit($buttonAction[4],"|")	;Looking for commandline arguements
			if $buttonsplit[0] > 1 Then						;If we have any
				DownloadAndRun($buttonsplit[1], "Retrieving Content...", $buttonsplit[2]) 	;Then use them
			Else
				DownloadAndRun($buttonAction[4], "Retrieving Content...", "")				;Otherwise, don't
			EndIf
		case '3'											;Read from the registry and execute - used for running the correct uninstall commands for Anyconnect, etc - guarentees path correctness.
			RunFromReg($buttonAction[4])					;No options on this one, Just do it.
	EndSwitch
EndFunc   ;==>DoButtonStuff

;Connectivity Test Function, rebuilt to allow for silent tests that the enduser doesn't see
Func ConnectivityTest($site, $bgcolor, $fgcolor, $retries = 500, $showwindow = 1)	;By default, show the window
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
	$ping = Ping($site, 900)	;ping the site.

	While @error ; if unable to contact the site selected, retry it, up to $retries times at 1s intervals, with a timeout of 900ms
		If $showwindow = 1 Then
			GUICtrlSetData($statusmessage, "Authorization Failed. Retrying.")
		EndIf
		If $failureCount < $retries Then	;Not out of attempts yet
			Sleep(1000)
			$failureCount = $failureCount + 1
		Else ; End the program after $retries attempts. The default amount is so high to give people time to log into the vpn.
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

	If $failureCount < $retries Then
		Return 0	;If ping test succeeded return 0
	Else
		Return 1 	;otherwise return 1
	EndIf
EndFunc   ;==>ConnectivityTest

Func LoadConfig($configurationFile)
	Return StringSplit(BinaryToString(InetRead($configurationFile, 17)), @CRLF, 1) ; Read Config file into an array of strings
EndFunc   ;==>LoadConfig

Func PrepareConfig($configfile, ByRef $tabs, ByRef $buttons, ByRef $abs_buttons) 	;Parse the Configfile, pass both of the
																					;parameter variables back to the main program
	$tabs = 0 ;total number of tabs
	$buttons = 0 ;total number of buttons
	Local $numberOfButtons = 0 ;local counter variable
	Local $i, $j ;local counter variables
	Local $configVersion

	For $i = 1 To $configfile[0] ;iterate through the array and determine how many lines are comments
		If StringLeft($configfile[$i], 1) <> '#' Then ;ignore comments and tab markers
			If StringLeft($configfile[$i], 1) = '^' Then	;If the first character is a ^
				$configVersion = Int(StringTrimLeft($configfile[$i], StringLen($configfile[$i])-1)) ;Set the version # of the config file ala DNS
			ElseIf StringLeft($configfile[$i], 1) = '@' Then	;Defines what to do with Tabs
				$tabs = $tabs + 1
				If $numberOfButtons > $buttons Then				;Makes sure we return the largest # of buttons an any tab.
					$buttons = $numberOfButtons
				EndIf
				$numberOfButtons = 0
			Else											;Found something that should be a button definition, may be mangled but we don't care about that yet
				For $j = 1 To $CmdLine[0]
					If StringLeft($configfile[$i], StringInStr($configfile[$i], ",") - 1) = $CmdLine[$j] Then ;ignore tenants that weren't passed to the module.
						$numberOfButtons = $numberOfButtons + 1
						$abs_buttons = $abs_buttons + 1 ;Track the number of buttons.
					EndIf
				Next
			EndIf
		EndIf
	Next
	If $numberOfButtons > $buttons Then					;Makes sure we return the largest # of buttons an any tab, even the last.
		$buttons = $numberOfButtons
	EndIf
	If $guiVersion < $configVersion Then			 ;Only adjust the GUI Size if the configfile has been updated - otherwise leave it alone (accomodates resize)
		$guiWidth=155+155+155+15
	EndIf
		If Floor($buttons / int($guiWidth / 155)) = ($buttons / int($guiWidth / 155)) Then ;Dynamically generate the GUI dimensions based on known button sizes. Width is currently standard
			$guiHeight = (Floor($buttons / int($guiWidth / 155)) * 30) + 10 + 110 + 30 ;If the number buttons are divisible by 5 use this formula
		Else
			$guiHeight = (Floor($buttons / int($guiWidth / 155)) * 30) + 40 + 110 + 30 ;If not, use this one.
		EndIf
Return $configVersion								;Function returns the version number of the configfile
EndFunc   ;==>PrepareConfig

Func DrawMainGUI()
	$hGUI = GUICreate("Support.com Workbench", $guiWidth, $guiHeight,$guiX,$guiY,BitOR($WS_CAPTION, $WS_MINIMIZEBOX, $WS_POPUP, $WS_SIZEBOX, $WS_SYSMENU)) ;Create Main GUI $WS_MAXIMIZEBOX,
	GUISetBkColor(0x000000) ;Set Background color
	GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked") ;Set Close event
	GUIRegisterMsg($WM_SIZE, "WBSize")
	GUISetOnEvent($GUI_EVENT_RESIZED, "WorkbenchResized")
	if $guiWidth > 460 Then
		$tabControl = GUICtrlCreateTab(0, 0, $guiWidth, $guiHeight - 120) ;Create the Tabbed Interface
	Else
		$tabControl = GUICtrlCreateTab(0, 0, $guiWidth, $guiHeight - 15) ;Create the Tabbed Interface
	EndIf
	GUICtrlSetResizing($tabControl, 2 + 4 + 32)
	Local $buttonCount = 1 ;Temporary Button count
	Local $realbuttonCount = 0 ;Actual Button count
	Local $tabCount = 1 ;Temporary Tab count

	For $i = 1 To $moduleConfig[0] ;Determine How to draw this
		If StringLeft($moduleConfig[$i], 1) <> '#' Then ;If the line is a comment, Discard it.
			$tempArray = StringSplit($moduleConfig[$i], ",") ;Read each line of the configfile into an array seperated by commas
			If $tempArray[1] = '@' Then ;If the line is a Tab declaration, set up a new tab with appropriate labels on the bottom
				For $j = 1 To $CmdLine[0]
					If $tempArray[2] = $CmdLine[$j] Then
						$TabArray[$tabCount] = GUICtrlCreateTabItem($tempArray[3]) ;New tab
						GUICtrlSetResizing(GUICtrlSetColor(GUICtrlCreateLabel("Internet Explorer Version: " & GetIEVersionFull(), 5, $guiHeight - 15), $textColor),2 + 4 + 32) ;Resizing options
						$tabCount = $tabCount + 1
						$buttonCount = 1
					EndIf
				Next
			EndIf
			If $tempArray[0] = '5' Then ;If the line is malformatted (or the version #), discard it
				For $j = 1 To $CmdLine[0] ;If the tenant isn't one of the tenants called by the login shell, discard it
					If $tempArray[1] = $CmdLine[$j] Then ;Found a match - To implement later, sup check
						$realbuttonCount = $realbuttonCount + 1
						$ButtonArray[$realbuttonCount] = GUICtrlCreateButton($tempArray[3], (Mod($buttonCount, int($guiWidth / 155)) * 155) + 10, (Ceiling($buttonCount / int($guiWidth / 155)) * 30), 150) ;Draws and assigns buttons
						GUICtrlSetResizing($ButtonArray[$realbuttonCount], 256+512+32+2)	;This tells the buttons not to move at all when you resize, since we redraw the GUI after a resize anyway.
						GUICtrlSetOnEvent($ButtonArray[$realbuttonCount], "DoButtonStuff") ;This function drives all the buttons.
						GUICtrlSetTip($ButtonArray[$realbuttonCount],$tempArray[5])
						$ButtonRefArray[$ButtonArray[$realbuttonCount]]=$moduleConfig[$i]	;Stores the button config info into $buttonRefArray at Array index ControlID.
						$buttonCount = $buttonCount + 1
					EndIf
				Next
			EndIf
		EndIf
	Next
	GUISetState(@SW_SHOW)
EndFunc   ;==>DrawMainGUI


Func DrawLogo()
	$picGUI = GUICreate("", 450, 105, ($guiWidth / 2) - 225, $guiHeight - 125, $WS_POPUP, BitOR($WS_EX_LAYERED, $WS_EX_MDICHILD), $hGUI) ;Create seperate GUI for transparent GIF
	GUICtrlCreatePic(@ScriptDir & "\logo.gif", 0, 0, 0, 0) ;Load Pic
	GUISetState(@SW_SHOW, $picGUI) ;Show Pic
EndFunc   ;==>DrawLogo

Func WorkbenchResized()			;This function fires after the window has finished being resized
    local $WB_Size = WinGetPos("[ACTIVE]")	;This function returns an array with 4 elements, X and Y coordinates in index 0/1 and Width/Height in index 2/3. Used for redrawing the GUI.
	$guiX=$WB_Size[0]					;Assign appropriate Global Variables to contain the positioning and size info
	$guiY=$WB_Size[1]					;"
	$guiWidth=$WB_Size[2]				;"
	$guiHeight=$WB_Size[3]				;"
	GUIDelete($picGUI)					;Remove the existing GUIs
	GUIDelete($hGUI)					;
	if $guiWidth < 165 Then				;Verify the GUI is wide enough to display at least 1 button.
		$guiWidth = 165					;If not, make it able to.
	EndIf
	if $guiWidth > 460 Then				;minimum Height calculations
		if $guiHeight < Int(($TopButtonsPerTab * 30 ) / int($guiwidth / 155)) + 175 Then	;Use this one if we're going to display the logo (3 or more buttons)
			$guiHeight = Int(($TopButtonsPerTab * 30 ) / int($guiwidth / 155)) + 175
		EndIf
	Else
		if $guiHeight < Int(($TopButtonsPerTab * 30 ) / int($guiwidth / 155)) + 60 Then		;Otherwise use this one.
			$guiHeight = Int(($TopButtonsPerTab * 30 ) / int($guiwidth / 155)) + 60
		EndIf
	EndIf
	DrawMainGUI()				;Redraw the GUI
	if $guiWidth > 460 Then		;Draw the logo if the program is wide enough.
		DrawLogo()
	EndIf
EndFunc	;==>WorkbenchResized

Func WBSize()						;This function triggers WHILE the window is being resized - hides the logo because it can expand out of the edge of the window.
	GuiSetState(@SW_HIDE, $picGUI)	;I don't like this solution, but it's the only one I've been able to find thus far.
EndFunc	;==>WBSize


;Function to kill the whole program
Func CLOSEClicked()
	Exit
EndFunc   ;==>CLOSEClicked

;Function to close out the Support Info window and re-enable the Control on the main window.
Func On_Support_Closed()
	GUIDelete($supportGUI)
	GUISetState(@SW_SHOW, $hGUI)
	GUISetState(@SW_SHOW, $picGUI)
EndFunc   ;==>On_Support_Closed
