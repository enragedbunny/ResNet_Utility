; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com [Please everyone add their contact email in this section]
; Date Edited: 2012.10.29

Func CreatePreferencesWindow()
	#Include <GuiConstantsEx.au3>

	GUICreate("Preferences", 710, 235)
	

	GUISetState(@SW_SHOW)

	While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE;Closes window if program is given
			Exit ;This Exit command is what actually makes the program exit.
	EndSwitch
WEnd

EndFunc