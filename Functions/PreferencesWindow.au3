; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com [Please everyone add their contact email in this section]
; Date Edited: 2012.10.29
Func CreatePreferencesWindow()
	#Include <GuiConstantsEx.au3>

	$PreferencesWindow = GUICreate("Preferences",300, 170)
	GUICtrlCreateLabel("Default Tech",5,5)
	GUICtrlCreateList("",85,3)
	GUICtrlCreateCheckbox("Autosave ticket when closing program?",5,30)
	GUICtrlCreateLabel("What color do you want the text to be? (ex 0x0000FF)?",5,55)
	$TextColor = IniRead("Functions\config\Config.ini","Preferences","Text_Color","NotFound")
	GUICtrlCreateInput($TextColor,5,78,100)
	GUICtrlCreateCheckbox("Automatically load tickets found in temp dir?",5,101)
	$btnOK = GUICtrlCreateButton("OK",5,126,80)
	$btnCancel = GUICtrlCreateButton("Cancel",90,126,80)

	GUISetState(@SW_SHOW)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE;Closes window if program is given
				GUIDelete($PreferencesWindow) ;This Exit command is what actually makes the program exit.
				Return
			Case $btnCancel
				GUIDelete($PreferencesWindow)
				Return
			Case $btnOK
				GUIDelete($PreferencesWindow) ;For now
				Return
		EndSwitch
	WEnd
EndFunc