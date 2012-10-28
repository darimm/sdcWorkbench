#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SDCWorkbench.ico
#AutoIt3Wrapper_Outfile=SDCWB_Uninstall.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Fileversion=0.1.0.0
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Obfuscator=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Version 0.1
;Written by Scott Stanley

DirRemove(EnvGet("PROGRAMFILES") & "\supportdotcom\SDCWorkbench",1)
DirRemove(@StartMenuCommonDir & "\Programs\supportdotcom\SDCWorkbench",1)
FileDelete(@DesktopDir & "\Support.com WorkBench.lnk")
MsgBox(0,"Removal Complete","Removal of the Support.com WorkBench is complete.")