; #FUNCTION# ;===============================================================================
;
; Name...........: _RunCMD
; Description ...: Takes CMD prompt commands as input and runs them in a cmd.exe window. has
;                  several optional parameters that allow some customization of how it runs
; Syntax.........: _RunCMD($sCommands, $iShowDialog, $iCloseWhenFinished)
; Parameters ....: $sCommands             - List of commands to run, separated by '|' (Pipe)
;                  $iShowDialog           - Input 1 or 0. 
;										   - 1 shows the command prompt
;										   - 0 hides the command prompt
;                  $iCloseWhenFinished    - Input 1 or 0.
;				   						   - 1 closes windows when command finishes
;										   - 0 leaves the window up when command is finished
; Return values .: None.
; Author ........: Johnny Keeton
; Modified.......: 2013.02.20
; Remarks .......: 
; Related .......: 
; Link ..........;
; Example .......; no
;
; ;==========================================================================================

Func _RunCMD($sCommands,$iShowDialog = 1,$iCloseWhenFinished = 0) ;Creates the function with options for data and optional flags to hide/show the window and leave open/close the window when finished.
	$asCMD = StringSplit($sCommands,'|') ;Split the input data, separating different strings by the | delimiter
	
	If $iCloseWhenFinished = 1 then ;Checks to see if we should close the window when finished
		$iCloseWhenFinished = "/c" ;Sets the variable to close the window when finished
	ElseIf $iCloseWhenFinished = 0 then ;Checks to see if we should leave the window open when finished
		$iCloseWhenFinished = "/k" ;Sets the variable to leave the window open when finished
	Else ;Error checking to be sure only correct flags are used
		Return ;[TODO] Need to return an errorcode of some kind. Currently just exits the function
	EndIf

	If IsArray($asCMD) then ;Checks that the StringSplit was successful and we are dealing with an array
		If $iShowDialog = 0 then
			For $iCount = 1 to $asCMD[0] ;Sets up the loop for the length of the array (length stored in $asCMD[0], because of StringSplit operation)
				RunWait('"' & @ComSpec & '" ' & $iCloseWhenFinished & " " & $asCMD[$iCount], @SystemDir, @SW_HIDE) ;Runs each command in the sequence, using the flags provided
			Next
		ElseIf $iShowDialog = 1 then
			For $iCount = 1 to $asCMD[0] ;Sets up the loop for the length of the array (length stored in $asCMD[0], because of StringSplit operation)
				RunWait('"' & @ComSpec & '" ' & $iCloseWhenFinished & " " & $asCMD[$iCount], @SystemDir, @SW_SHOW) ;Runs each command in the sequence, using the flags provided
			Next
		Else
			;Need to return error [TODO]
			return
		EndIf
	EndIf
EndFunc