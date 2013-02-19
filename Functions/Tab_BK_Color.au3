; Program: Change Tab background Color
; Author: Alex Burgy-Vanhoose
; Other Contributors: Johnny Keeton
; Contact: johnny.keeton@gmail.com
; Date Edited: 2013.02.18

Func _GUICtrlTab_SetBkColor($hWnd, $hSysTab32, $sBkColor)

    Local $aTabPos = ControlGetPos($hWnd, "", $hSysTab32)
    Local $aTab_Rect = _GUICtrlTab_GetItemRect($hSysTab32, -1)

    GUICtrlCreateLabel("", $aTabPos[0]+2, $aTabPos[1]+$aTab_Rect[3]+4, $aTabPos[2]-6, $aTabPos[3]-$aTab_Rect[3]-7)
    GUICtrlSetBkColor(-1, $sBkColor)
    GUICtrlSetState(-1, $GUI_DISABLE)
EndFunc