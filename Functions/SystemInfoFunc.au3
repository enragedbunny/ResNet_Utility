; File: SystemInfoFunc.au3
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com 
; Date Edited: 2012.10.11
; Purpose: Provide WMI access to collect system information
; Compatability: Windows Vista/7 (probably XP, will work for Windows 8)

Func GET_Device_Manager_Errors($objWMI) ;retrieve device manager errors (not implemented yet) will not return errors, but instead return if there are errors or not
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_PnPEntity
	If IsObj($objItems) Then
		For $objItem In $objItems
			;report that errors exist
		Next
	EndIf
EndFunc
Func GET_Manufacturer_and_Model($objWMI) ; Takes WMI object as input, returns Local System's Manufacturer and Model
	$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_ComputerSystem
	If IsObj($objItems) Then
		For $objItem In $objItems
			Return $objItem.Manufacturer & "|" & $objItem.Model ;return manufacter and model for local system. This is formatted "Manufacturer|Model" and my need further procesing if used in another program
		Next
	EndIf
EndFunc

Func GET_Serial_Number($objWMI) ; Takes WMI object as input, returns Local System's Serial Number
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_BIOS
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
		return String(Int($Counter / (1024^3))) & " GB" ;Converts B to GB, formats and returns total RAM
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
			If $objItem.BuildNumber >= 8000 Then ;Checks for Windows 8
				Return "Win 8" & "|" & String(Int($objItem.BuildNumber)-7600) ;Returns OS and strips service pack from build number
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
	local $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", 0x10 + 0x20) ;query WMI to retrieve Win32_NetworkAdapter
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
		If $WifiSettings = "" Then ; If there is no wifi Adapter, loads default data
			$WifiSettings = "No Adapter|##:##:##:##:##:##|" ;Indicates there is no adapter
		EndIf
		If $WiredSettings = "" Then ;If there is no wired adapter, loads default data
			$WiredSettings = "No Adapter|##:##:##:##:##:##" ;Indicates there is no adapter
		EndIf
		Return $WifiSettings & $WiredSettings ; returns wireless and wired description and MAC Address in the format "WirelessDescription|MACAddress|WiredDescription|MACAddress".
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