; File: SystemInfoFunc.au3
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com 
; Date Edited: 2012.11.2
; Purpose: Provide WMI access to collect system information
; Compatability: Windows Vista/7 (probably XP, will work for Windows 8)

<<<<<<< HEAD
Func GET_Device_Manager_Errors($objWMI) ;retrieve device manager errors (Coded, NOT tested) will not return errors, but instead tell you if there are any
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_PnPEntity
	If IsObj($objItems) Then
		$DMErrors = "Errors Exist"
		For $objItem In $objItems
			$DMErrors = $SmErrors & "|" & $objItem.Caption
		Next
	Else
		$DMErrors = "None"
	EndIf
	Local $DMErrArray = StringSplit($DMErrors,"|")
	Return $DMErrArray
EndFunc

=======
Func GET_Device_Manager_Errors($objWMI)
   	;Check device manager for errorss (Not implemented, but do not delete this)
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0", "WQL", 0x10 + 0x20)
	If IsObj($objItems) Then
		For $objItem In $objItems
			;report that errors exist
		Next
	EndIf
EndFunc
>>>>>>> Updated SysInfoFunc and modified main utility to make changes work
Func GET_Manufacturer_and_Model($objWMI) ; Takes WMI object as input, returns Local System's Manufacturer and Model
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_ComputerSystem
	If IsObj($objItems) Then
		For $objItem In $objItems
			Return $objItem.Manufacturer & "|" & $objItem.Model ;return manufacter and model for local system. This is formatted "Manufacturer|Model" and my need further procesing if used in another program
		Next
	EndIf
EndFunc

Func GET_Serial_Number($objWMI) ; Takes WMI object as input, returns Local System's Serial Number
<<<<<<< HEAD
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_BIOS
=======
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", 0x10 + 0x20)
>>>>>>> Updated SysInfoFunc and modified main utility to make changes work
	If IsObj($objItems) Then
		For $objItem In $objItems
			Return $objItem.SerialNumber ;return serial number of local machine
		Next
	EndIf
EndFunc

Func GET_HDD_Total_and_Free($objWMI) ; Takes WMI object as input, returns Local System's Hard Drive Total space and Free space
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_LogicalDisk
	If IsObj($objItems) Then
		local $Counter = 0,$Counter2 = 0 ;Declare and initialize variables used for counters
		For $objItem In $objItems ; This loop adds up each partition/hard drive to combine all sizes and free space together
			$Counter = $Counter + Int($objItem.Size) ;Add up total disk space
			$Counter2 = $Counter2 + Int($objItem.FreeSpace) ;Add up free space
		Next
		return String(Int($Counter2 / (1024^3))) & " / " & String(Int($Counter / (1024^3))) & " GB" ;Converts HDD size and free space from B to GB and formats in a logical way
	EndIf
EndFunc

Func GET_Total_RAM($objWMI) ; Takes WMI object as input, returns Local System's total RAM
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_PhysicalMemory
	If IsObj($objItems) Then
		local $Counter = 0
		For $objItem In $objItems ; This loop combines each RAM slot
			$Counter = $Counter + Int($objItem.Capacity)
		Next
<<<<<<< HEAD
		return String(Int($Counter / (1024^3))) & " GB" ;Converts B to GB, formats and returns total RAM
=======
		return String(Int($Counter / (1024^3))) & " GB" ;Converts B to GB, formats and returns the result
>>>>>>> Updated SysInfoFunc and modified main utility to make changes work
	EndIf
EndFunc

Func GET_System_Architecture($objWMI) ; Takes WMI object as input, returns Local System's Architecture (32 or 64 bit)
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_Processor", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_Processor
	If IsObj($objItems) Then
		For $objItem In $objItems
			If $objItem.AddressWidth == 64 Then ;Checks to see if 64 bit or not
				Return "64-bit" ;Returns 64-bit
			Else ;only other option is 32-bit
				Return "32-bit" ;Returns 32-bit
			EndIf
		Next
	EndIf
EndFunc

Func GET_OS_and_Service_Pack($objWMI) ; Takes WMI object as input, returns Local System's OS version and Service pack, if any
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_OperatingSystem
	If IsObj($objItems) Then
		For $objItem In $objItems
			If $objItem.BuildNumber >= 9200 Then ;Checks for Windows 8
				Return "Win 8" & "|" & String(Int($objItem.BuildNumber)-9200) ;Returns OS and strips service pack from build number
			ElseIf $objItem.BuildNumber >= 7600 Then ;Checks for Win 7 OS
				Return "Win 7" & "|" & String(Int($objItem.BuildNumber)-7600) ;Returns OS and Strips service pack from build number
			ElseIf	$objItem.BuildNumber >= 6000 Then ;Checks for Windows Vista os
				Return "Win Vista" & "|" & String(Int($objItem.BuildNumber)-6000) ;Returns OS and strips service pack from build number
			EndIf
		Next
	EndIf
EndFunc

Func GET_Ethernet_and_Wireless($objWMI) ; Takes WMI object as input, returns Local System's Wired and Wireless device description and MAC Address
	local $WifiSettings,$WiredSettings,$CombinedResults
<<<<<<< HEAD
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_NetworkAdapter
=======
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", 0x10 + 0x20) ; Wifi/Wireless MAC Addresses and IPs
>>>>>>> Updated SysInfoFunc and modified main utility to make changes work
	If IsObj($objItems) Then
		For $objItem In $objItems
			;The following line does some filtering. It makes sure the entry has a MAC address, and is not WiMAX, Miniport, or bluetooth.
			If Not($objItem.MACAddress = "None") And StringInStr(String($objItem.Description),"WiMAX") = 0 And StringInStr(String($objItem.Description),"Miniport") = 0 And StringInStr(String($objItem.Description),"Bluetooth") = 0 Then
				If Not(StringInStr($objItem.NetConnectionID,"Wireless") = 0) or Not(StringInStr($objItem.NetConnectionID,"Wi-Fi") = 0) Then ;This automatically detects wireless adapters and collects info
					$WifiSettings = $objItem.Description & "|" & $objItem.MACAddress & "|"
				ElseIf Not(StringInStr($objItem.NetConnectionID,"Local") = 0) or Not(StringInStr($objItem.NetConnectionID,"Ethernet") = 0) Then ;This automatically detects wired adapters and collects info
					$WiredSettings = $objItem.Description & "|" & $objItem.MACAddress ;Formats network information as "Description|MACAddress" may need further processing if used in other programs.
				EndIf
			EndIf
		Next
<<<<<<< HEAD
		If $WifiSettings = "" Then ; If there is no wifi Adapter, loads default data
			$WifiSettings = "No Adapter|##:##:##:##:##:##|" ;Indicates there is no adapter
		EndIf
		If $WiredSettings = "" Then ;If there is no wired adapter, loads default data
			$WiredSettings = "No Adapter|##:##:##:##:##:##" ;Indicates there is no adapter
		EndIf
		Return $WifiSettings & $WiredSettings ; returns wireless and wired description and MAC Address in the format "WirelessDescription|MACAddress|WiredDescription|MACAddress".
=======
		Return $WifiSettings & "|" & $WiredSettings
>>>>>>> Updated SysInfoFunc and modified main utility to make changes work
	EndIf
EndFunc

Func GET_Workgroup($objWMI) ; Takes WMI object as input, returns Local System's Workgroup
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_ComputerSystem
	If IsObj($objItems) Then ;checks to see that the query returned usable results
		For $objItem In $objItems ;for each item in the returned query
			return $objItem.Workgroup ; Returns the Workgroup
		Next
	EndIf
EndFunc