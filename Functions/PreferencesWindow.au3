; Program: Preferences Window
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose
; Contact: johnny.keeton@gmail.com 
; Date Edited: 2012.11.2

Func CreatePreferencesWindow()
	#Include <GuiConstantsEx.au3>
	local $ConfigPath = "Config.ini"

	$PreferencesWindow = GUICreate("Preferences",300, 170)

	GUICtrlCreateLabel("Default Tech",5,5)
	GUICtrlCreateCombo("Select Name",85,3,140)
	Local $DefaultTech = iniRead($ConfigPath,"Preferences","Default_Tech","Select Name")
	GUICtrlSetData(-1, "Johnny Keeton|Logan Wallen|Zack Smith|CJ Highley|Adam Weintraut|Alex Burgy-Vanhoose|Josh Back|Kia Wallen|Courtney Isaacs|Lisa Moore",$DefaultTech)
	
	GUICtrlCreateCheckbox("Autosave ticket when closing program?",5,30)
	Local $AutoSave = IniRead($ConfigPath,"Preferences","AutoSave_On_Close","NotFound")
	
	If Not($AutoSave = "NotFound") Then
		If $AutoSave = 1 Then
			GUICtrlSetState(-1,$GUI_CHECKED)
		ElseIf $AutoSave = 0 Then 
			GUICtrlSetState(-1,$GUI_UNCHECKED)
		EndIf
	EndIf

	GUICtrlCreateLabel("What color do you want the text to be? (ex 0x0000FF)?",5,55)
	Local $TextColor = IniRead($ConfigPath,"Preferences","Text_Color","NotFound")
	GUICtrlCreateInput($TextColor,5,78,100)
	
	GUICtrlCreateCheckbox("Automatically load tickets found in temp dir?",5,101)
	Local $AutoLoad = IniRead($ConfigPath,"Preferences","AutoLoad_Ticket","NotFound")

	If Not($AutoLoad = "NotFound") Then
		If $AutoLoad = 1 Then
			GUICtrlSetState(-1,$GUI_CHECKED)
		ElseIf $AutoLoad = 0 Then
			GUICtrlSetState(-1,$GUI_UNCHECKED)
		EndIf
	EndIf
	$btnOK = GUICtrlCreateButton("OK",5,126,80)
	$btnCancel = GUICtrlCreateButton("Cancel",90,126,80)

	GUISetState(@SW_SHOW)

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE;Closes window if program is given
				GUIDelete($PreferencesWindow)
				Return
			Case $btnCancel
				GUIDelete($PreferencesWindow)
				Return
			Case $btnOK
				GUIDelete($PreferencesWindow)
				Return
		EndSwitch
	WEnd
EndFunc