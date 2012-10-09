; Program: SMART Functions
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com [Please everyone add their contact email in this section]
; Date Edited: 2012.10.07
; NOTE: I (JK) modified this code from another source from the internet. If the owner recognizes this code and does not wish it to be
;       part of this project, please let me know and I will promptly remove it. 

Func Initialize_SMART() ; This function calls other functions and serves to parse results from errors
	Local $drive = "C:" ;Drive to get SMART for
	Local $smartData = _GetSmartData($drive) ;This runs function _GetSmartData with the $drive as the drive to check
	If $smartData <> -1 then ; If _GetSmartData did NOT return an error (Error returns value of -1)
		_arrayDisplay($SmartData,"S.M.A.R.T. Data for Drive " & $drive) ; Call function _arrayDisplay to display the SMART Data
	Else ; If there was an error
		Msgbox(0,"",$drive & " may not be S.M.A.R.T. Capable") ; Generic message for receiving an error for the drive
	EndIf
EndFunc

Func _GetSmartData($vDrive  = "C:") ;optional Parameter $vDrive is = C: by default

	If DriveStatus ( $vDrive ) = "INVALID" then ;Checks to see if SMART can be determined for the drive
		Return -1 ; Exits the function and sets the status as "-1" to indicate some error occurred
	EndIf

	If StringLeft($vDrive,1) <> '"' then $vDrive = '"' & $vDrive ;Checks (and ensures) to see if the drive begins with a quotation
	If StringRight($vDrive,1) <> '"' then $vDrive &= '"' ;Checks (and ensures) to see if the drive ends with a quotation

	Local $iCnt, $iCheck, $vPartition,$vDeviceID,$vPNPID ; Declares local variables needed.

	$vPartition = _LogicalToPartition ($vDrive) ; Finds Partition ID from Logical drive ID (C:)
	If $vPartition = -1 then Return -1 ; If there was an error, stops function and returns error -1

	$vDeviceID = _PartitionToPhysicalDriveID($vPartition) ; Finds Drive ID from Partition ID
	If $vDeviceID = -1 then Return -1 ; If there was an error, stops function and returns error -1

	$vPNPID = _PNPIDFromPhysicalDriveID($vDeviceID) ; Finds PlugNPlay ID from Physical Drive ID
	If $vPNPID = -1 then Return -1 ; If there was an error, stops function and returns error -1

	local $strComputer = "."
	local $objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\WMI") ; Gets WMI object from local system
	local $oDict = ObjCreate("Scripting.Dictionary") ; Creates new Object (Dictionary)
	local Const $wbemFlagReturnImmediately = 0x10 ; Flag for SQL Query on WMI
	local Const $wbemFlagForwardOnly = 0x20 ; Flag for SQL Query on WMI
		

	local $colItems1 = $objWMIService.ExecQuery("SELECT * FROM MSStorageDriver_ATAPISmartData", "WQL", _ ; Saves information from WMI to be used later
										  $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	local $colItems2 = $objWMIService.ExecQuery("SELECT * FROM MSStorageDriver_FailurePredictThresholds", "WQL", _ ; Saves information from WMI to be used later
										  $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	local $colItems3 = $objWMIService.ExecQuery("SELECT * FROM MSStorageDriver_FailurePredictStatus", "WQL", _ ; Saves information from WMI to be used later
										  $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

	_CreateDict($oDict) ; Create Labels for SMART data entries (NEED to figure out how to do this with local variables.)

	ConsoleWrite ($strComputer&@CR) ;@CR is a carriage return (line return)

	local $iCnt = 1 ; Counts the number of devices that are not the target one
	For $objItem In $colItems1 ; loop through each found item from the query on ATAPISmartData from WMI
		If StringLeft($objItem.InstanceName,StringLen($vPNPID)) = String($vPNPID) then ;The StringLen is limiting the comparison of the data. If the PNPID matches the device, the statement is true
			ConsoleWrite ("Active: " & $objItem.Active&@CR)
			ConsoleWrite ("Checksum: " & $objItem.Checksum&@CR)
			ConsoleWrite ("ErrorLogCapability: " & $objItem.ErrorLogCapability&@CR)
			ConsoleWrite ("ExtendedPollTimeInMinutes: " & $objItem.ExtendedPollTimeInMinutes&@CR)
			ConsoleWrite ("InstanceName: " & $objItem.InstanceName&@CR)
			ConsoleWrite ("Length: " & $objItem.Length&@CR)
			ConsoleWrite ("OfflineCollectCapability: " & $objItem.OfflineCollectCapability&@CR)
			ConsoleWrite ("OfflineCollectionStatus: " & $objItem.OfflineCollectionStatus&@CR)
			local $strReserved = $objItem.Reserved
			ConsoleWrite ("Reserved: " & _ArrayToString($strReserved,",")&@CR)
			ConsoleWrite ("SelfTestStatus: " & $objItem.SelfTestStatus&@CR)
			ConsoleWrite ("ShortPollTimeInMinutes: " & $objItem.ShortPollTimeInMinutes&@CR)
			ConsoleWrite ("SmartCapability: " & $objItem.SmartCapability&@CR)
			ConsoleWrite ("TotalTime: " & $objItem.TotalTime&@CR)
			local $strVendorSpecific = $objItem.VendorSpecific
			ConsoleWrite ("VendorSpecific: " & _ArrayToString($strVendorSpecific,",")&@CR)
			ConsoleWrite ("VendorSpecific2: " & $objItem.VendorSpecific2&@CR)
			ConsoleWrite ("VendorSpecific3: " & $objItem.VendorSpecific3&@CR)
			local $strVendorSpecific4 = $objItem.VendorSpecific4
			ConsoleWrite ("VendorSpecific4: " & _ArrayToString($strVendorSpecific4,",")&@CR)
		Else

			$iCnt +=1 ; increases the counter
		EndIf

	Next

	local $iCheck = 1 ;declare and initialize local variable
	For $objItem In $colItems2 ;Loop through each item found in the second WMI query
		If $iCheck = $iCnt then ;compares the previous counter to this one
			local $strVendorSpecific2 = $objItem.VendorSpecific
			ConsoleWrite ("FailurePredictThresholds: " & _ArrayToString($strVendorSpecific2,",")&@CR)
		Else
			$iCheck +=1 ; if they didn't match, increase the counter
		EndIf
	Next

	$iCheck = 1 ;Reinitializes the variable
	For $objItem In $colItems3 ; loops through all items found in the final WMI query
		If $iCheck = $iCnt then
			local $strVendorSpecific3 = $objItem.PredictFailure
			ConsoleWrite ("PredictFailure: " & $strVendorSpecific3&@CR)
			local $strVendorSpecific3 = $objItem.Reason
			ConsoleWrite ("PredictReason: " & $strVendorSpecific4&@CR)
		Else
			$iCheck +=1
		EndIf
	Next

	If NOT IsArray($strVendorSpecific4) then Return -1 ;exit the function, not a Smart capable drive
	If NOT IsArray($strVendorSpecific2) then Return -1 ;exit the function, not a Smart capable drive

	local $Status [22] ;declaring $status 
	For $i = 1 to 21
		$Status[$i] = "NotOk" ;default all status to "NotOK"
	Next

	;if a status passes then Change it to OK
	If $strVendorSpecific[5] >= $strVendorSpecific2[3]  then $Status[1] = "OK"
	If $strVendorSpecific[17] >= $strVendorSpecific2[15] then $Status[2] = "OK"
	If $strVendorSpecific[29] >= $strVendorSpecific2[27] then $Status[3] = "OK"
	If $strVendorSpecific[41] >= $strVendorSpecific2[39] then $Status[4] = "OK"
	If $strVendorSpecific[53] >= $strVendorSpecific2[51] then $Status[5] = "OK"
	If $strVendorSpecific[65] >= $strVendorSpecific2[63] then $Status[6] = "OK"
	If $strVendorSpecific[77] >= $strVendorSpecific2[75] then $Status[7] = "OK"
	If $strVendorSpecific[89] >= $strVendorSpecific2[87] then $Status[8] = "OK"
	If $strVendorSpecific[101] >= $strVendorSpecific2[99] then $Status[9] = "OK"
	If $strVendorSpecific[113] >= $strVendorSpecific2[111] then $Status[10] = "OK"
	If $strVendorSpecific[125] >= $strVendorSpecific2[123] then $Status[11] = "OK"
	If $strVendorSpecific[137] >= $strVendorSpecific2[135] then $Status[12] = "OK"
	If $strVendorSpecific[149] >= $strVendorSpecific2[147] then $Status[13] = "OK"
	If $strVendorSpecific[161] >= $strVendorSpecific2[159] then $Status[14] = "OK"
	If $strVendorSpecific[173] >= $strVendorSpecific2[171] then $Status[15] = "OK"
	If $strVendorSpecific[185] >= $strVendorSpecific2[183] then $Status[16] = "OK"
	If $strVendorSpecific[197] >= $strVendorSpecific2[195] then $Status[17] = "OK"
	If $strVendorSpecific[206] >= $strVendorSpecific2[204] then $Status[18] = "OK"
	If $strVendorSpecific[218] >= $strVendorSpecific2[216] then $Status[19] = "OK"
	If $strVendorSpecific[230] >= $strVendorSpecific2[228] then $Status[20] = "OK"
	If $strVendorSpecific[242] >= $strVendorSpecific2[240] then $Status[21] = "OK"


	local $aSmartData[1][8] = [["ID","Attribute","Type","Flag","Threshold","Value","Worst","Status"]]
	$iCnt=1 ;reinitialize variable
	For $x = 2 to 242 Step 12
		If $strVendorSpecific[$x] <> 0 then ;0 id is not valid for this Smart Device
			local $NextRow = Ubound($aSmartData)
			Redim $aSmartData[$NextRow +1][8]
				$aSmartData[$NextRow][0] = $strVendorSpecific[$x];ID
				$aSmartData[$NextRow][1] = $oDict.Item($strVendorSpecific[$x]);Attribute
				If $aSmartData[$NextRow][1] = "" then $aSmartData[$NextRow][1] = "Unknown SMART Attribute"

				If Mod($strVendorSpecific[$x +1],2) then ;Type If odd number then it is a pre-failure value
					$aSmartData[$NextRow][2] = "Pre-Failure"
				Else
					$aSmartData[$NextRow][2] = "Advisory"
				EndIf
				
				$aSmartData[$NextRow][3] = $strVendorSpecific[$x +1];Flag
				$aSmartData[$NextRow][4] = $strVendorSpecific2[$x +1];Threshold
				$aSmartData[$NextRow][5] = $strVendorSpecific[$x +3];Value
				$aSmartData[$NextRow][6] = $strVendorSpecific[$x +4];Worst
				$aSmartData[$NextRow][7] = $Status[$iCnt];Status
		EndIf
		$iCnt+=1
	Next
	Return 	$aSmartData
EndFunc

Func _PartitionToPhysicalDriveID($vPart)
	local $objWMIService,$colItems,$vPdisk,$strDeviceID,$strComputer
	local Const $wbemFlagReturnImmediately = 0x10 ; Flag for SQL Query on WMI
	local Const $wbemFlagForwardOnly = 0x20 ; Flag for SQL Query on WMI
	$strComputer = "."
	
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_DiskDriveToDiskPartition", "WQL", _
                                          $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If IsObj($colItems) then
		For $objItem In $colItems
			If $objItem.Dependent == $vPart then
				$vPdisk = $objItem.Antecedent ; saves data from query to string
				$vPdisk = StringReplace(StringTrimLeft($vPdisk,Stringinstr($vPdisk,"=")),'"',"") ; formats string
				$strDeviceID = StringReplace($vPdisk, "\\", "\") ; more formatting
				Return $strDeviceID ; sends physical drive ID 
			Endif
		Next
	Endif
	Return -1 ; Returns error if it was unable to collect information
EndFunc

Func _PNPIDFromPhysicalDriveID($vDriveID)
	local $objWMIService,$colItems,$strComputer ;Declare variables
	local Const $wbemFlagReturnImmediately = 0x10 ; Flag for SQL Query on WMI
	local Const $wbemFlagForwardOnly = 0x20 ; Flag for SQL Query on WMI
	$strComputer = "."
	
	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive", "WQL", _
                                          $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If IsObj($colItems) then
		For $objItem In $colItems
			If $objItem.DeviceID == $vDriveID then
				Return $objItem.PNPDeviceID ; Plug N Play ID
			EndIf
		Next
	Endif
	Return -1 ; Returns error if it was unable to collect information
EndFunc

Func _LogicalToPartition ($vDriveLet)
	
	local $objWMIService,$colItems,$vAntecedent,$vPartition,$strComputer ; Declare variables
	local Const $wbemFlagReturnImmediately = 0x10 ; Flag for SQL Query on WMI
	local Const $wbemFlagForwardOnly = 0x20 ; Flag for SQL Query on WMI
	$strComputer = "."

	$objWMIService = ObjGet("winmgmts:\\" & $strComputer & "\root\CIMV2")
	$colItems = $objWMIService.ExecQuery("SELECT Dependent,Antecedent FROM Win32_LogicalDiskToPartition", "WQL", _
											  $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	If IsObj($colItems) then
		For $objItem In $colItems
			If StringRight($objItem.Dependent,4) == $vDriveLet then
				$vAntecedent=$objItem.Antecedent
				Return $vAntecedent ; Partition 
			EndIf
		Next
	Endif
	Return -1 ; Returns error if it was unable to collect information
EndFunc

Func _CreateDict(ByRef $oDict) 
	$oDict.Add(1,"Read Error Rate")
	$oDict.Add(2,"Throughput Performance")
	$oDict.Add(3,"Spin-Up Time")
	$oDict.Add(4,"Start/Stop Count")
	$oDict.Add(5,"Reallocated Sectors Count")
	$oDict.Add(6,"Read Channel Margin")
	$oDict.Add(7,"Seek Error Rate Rate")
	$oDict.Add(8,"Seek Time Performance")
	$oDict.Add(9,"Power-On Hours (POH)")
	$oDict.Add(10,"Spin Retry Count")
	$oDict.Add(11,"Recalibration Retries")
	$oDict.Add(12,"Device Power Cycle Count")
	$oDict.Add(13,"Soft Read Error Rate")
	$oDict.Add(191,"G-Sense Error Rate Frequency")
	$oDict.Add(192,"Power-Off Park Count")
	$oDict.Add(193,"Load/Unload Cycle")
	$oDict.Add(194,"HDA Temperature")
	$oDict.Add(195,"ECC Corrected Count")
	$oDict.Add(196,"Reallocated Event Count")
	$oDict.Add(197,"Current Pending Sector Count")
	$oDict.Add(198,"Uncorrectable Sector Count")
	$oDict.Add(199,"UltraDMA CRC Error Count")
	$oDict.Add(200,"Write Error Rate")
	$oDict.Add(201,"Soft Read Error Rate")
	$oDict.Add(202,"Address Mark Errors Frequency")
	$oDict.Add(203,"ECC errors (Maxtor: ECC Errors)")
	$oDict.Add(204, "Soft ECC Correction")
	$oDict.Add(205, "Thermal Asperity Rate (TAR)")
	$oDict.Add(206, "Flying Height")
	$oDict.Add(207, "Spin High Current")
	$oDict.Add(208, "Spin Buzz")
	$oDict.Add(209, "Offline Seek")
	$oDict.Add(220, "Disk Shift")
	$oDict.Add(221, "G-Sense Error Rate")
	$oDict.Add(222, "Loaded Hours")
	$oDict.Add(223, "Load/Unload Retry Count")
	$oDict.Add(224, "Load Friction")
	$oDict.Add(225, "/Unload Cycle Count")
	$oDict.Add(226, "Load 'In'-time")
	$oDict.Add(227, "Torque Amplification Count")
	$oDict.Add(228, "Power-Off Retract Cycle")
	$oDict.Add(230, "GMR Head Amplitude")
	$oDict.Add(231, "Temperature")
	$oDict.Add(240, "Head Flying Hours")
	$oDict.Add(250, "Read Error Retry Rate")
EndFunc