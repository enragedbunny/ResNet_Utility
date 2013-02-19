; Program: About Help manager
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose
; Contact: johnny.keeton@gmail.com 
; Date Edited: 2012.11.2

Func CreateAboutWindow($ProgramTitle,$Version,$ReleaseDate)
	#Include <GuiConstantsEx.au3>

	$AboutWindow = GUICreate("About",170, 80)
	GUICtrlCreateLabel("Program Name: " & $ProgramTitle,5,5)
	GUICtrlCreateLabel("Version: " & $Version,5,30)
	GUICtrlCreateLabel("Release Date: " & $ReleaseDate,5,55)

	GUISetState(@SW_SHOW)
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE;Closes window if program is given
				GUIDelete($AboutWindow) ;This closes the window.
				Return ;Exit the loop
		EndSwitch
	WEnd

EndFunc
Func CreateHelpWindow($FileName)
	Run("notepad.exe " & $Filename,"",@SW_MAXIMIZE)
EndFunc