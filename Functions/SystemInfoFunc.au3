; File: SystemInfoFunc.au3
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com 
; Date Edited: 2012.10.06
; Purpose: Provide WMI access to collect system information
; Notes:	1. Only works for windows system
;			2. UsesGUICtrlSetData to set the value of a field (may not be best way to do this)
;			3. It aims to collect as much (needed) information from each WMI object as possible
;Uses the following variables for setting label data:
;	$Bran = Device (computer) manufacturer
;	$Mode = Device (computer) Model
;	$Seri = Serial Number (as reported in BIOS
;	$HsHf = HDD size and free space, formatted as (###/### GB)
;	$Memo = Total RAM
;	$SyTy = System Type (Architecture)
;	$OSEd = Operation system (currently only detects Win 7 and Win Vista)
;	$SePa = Service Pack (Win 7 and Vista)
;	$WiBr = Wireless Brand/Manufacturer
;	$WiMa = Wireless MAC Address
;	$WrBr = Wired Brand/Manufacturer
;	$WrMa = Wired MAC Address
;	$Workgroup = Current Workgroup (used to check agains "RESNET" to see if it needs changed)
;	$Counter1,$Counter2 = used in several loops as counter variables, should be cleared after each loop

Func systemInfo()
   ;Create connection to WMI
   $objWMI = ObjGet("winmgmts:\\localhost\root\CIMV2")

   ;Get Manufacturer and Model
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
      For $objItem In $objItems
   	  GUICtrlSetData($Bran,$objItem.Manufacturer)
   	  GUICtrlSetData($Mode,$objItem.Model)
      Next
   EndIf

   ;Check device manager for errorss
   ;$objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0", "WQL", 0x10 + 0x20)
   ;If IsObj($objItems) Then
   ;   For $objItem In $objItems
   ;	  ;report that errors exist
   ;   Next
   ;EndIf

   ; Get Serial Number
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_BIOS", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
      For $objItem In $objItems
   	  GUICtrlSetData($Seri,$objItem.SerialNumber)
      Next
   EndIf

   ; Hard Drive Total and Free space
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
	  $Counter1 = 0
	  $Counter2 = 0
      For $objItem In $objItems ; This loop adds up each partition/hard drive to combine all sizes and free space together
		 $Counter1 = $Counter1 + Int($objItem.Size)
		 $Counter2 = $Counter2 + Int($objItem.FreeSpace)
      Next
   ;Converts HDD size and free space from B to GB and formats in a logical way
   GUICtrlSetData($HsHf,String(Int($Counter2 / (1024^3))) & " / " & String(Int($Counter1 / (1024^3))) & " GB")
   $Counter1 = ""
   $Counter2 = ""
   EndIf

   ; RAM information
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
	  $Counter1 = 0
      For $objItem In $objItems ; This loop combines each RAM slot
		 $Counter1 = $Counter1 + Int($objItem.Capacity)
      Next
   ;Converts B to GB and formats the result
   GUICtrlSetData($Memo,String(Int($Counter1 / (1024^3))) & " GB")
   $Counter1 = ""
   EndIf

   ; Architecture
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_Processor", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
      For $objItem In $objItems
		 If $objItem.AddressWidth == 64 Then ; Checks to see if 64 bit or not
			GUICtrlSetData($SyTy,"64-bit")
		 Else
			GUICtrlSetData($SyTy,"32-bit")
		 EndIf
      Next
   EndIf

   ; OS Version
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
      For $objItem In $objItems
		 If $objItem.BuildNumber >= 7600 Then ; Checks for Win 7 OS (or higher)
			GUICtrlSetData($OSEd,"Win 7")
			GUICtrlSetData($SePa,Int($objItem.BuildNumber)-7600) ;Service pack 1 would be 7601, so this trims down to service pack number
		 ElseIf	$objItem.BuildNumber >= 6000 Then
			GUICtrlSetData($OSEd,"Vista")
			GUICtrlSetData($SePa,Int($objItem.BuildNumber)-6000) ;Same here, just figuring out service pack
		 EndIf
      Next
   EndIf

   ; Wifi/Wireless MAC Addresses and IPs
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
      For $objItem In $objItems
		 ;The following line does some filtering. It makes sure the entry has a MAC address, and is not WiMAX, Miniport, or bluetooth.
		 If Not($objItem.MACAddress = "None") And StringInStr(String($objItem.Description),"WiMAX") = 0 And StringInStr(String($objItem.Description),"Miniport") = 0 And StringInStr(String($objItem.Description),"Bluetooth") = 0 Then
			;$Counter1 = ""  ; Not used right now, will be used to figure out IPAddress
			If Not(StringInStr($objItem.NetConnectionID,"Wireless") = 0) Then ; This automatically detects wireless adapters and collects info
			   GUICtrlSetData($WiBr, $objItem.Description)
			   GUICtrlSetData($WiMa,$objItem.MACAddress)
			   ;$Counter1 = $objItem.IPAddress ;Working on IPAddress detection
			   ;GUICtrlSetData($IPAd,$Counter1.substring(0,6)
			ElseIf Not(StringInStr($objItem.NetConnectionID,"Local") = 0) Then ; This automatically detects wired adapters and collects info
			   GUICtrlSetData($WrBr,$objItem.Description)
			   GUICtrlSetData($WrMa,$objItem.MACAddress)
			EndIf
		 EndIf
	  Next
   EndIf

   ; Workgroup
   $objItems = $objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", 0x10 + 0x20)
   If IsObj($objItems) Then
	   For $objItem In $objItems
		 If Not($objItem.Workgroup = "RESNET") Then ; This will alert user if workgroup is not RESNET so they can change it
			GUICtrlSetData($Workgroup,"Change Workgroup")
		 EndIf
	   Next
	EndIf
EndFunc







; The following code is a prototype for breaking the above code into separate functions that do not rely on the GUI variables to work correctly
; If you don't know what you are doing, don't edit any code below this point. Thanks (JK)
Func GET_Manufacturer_and_Model() ; Takes WMI object as input, returns Local System's Manufacturer and Model
	
EndFunc

Func GET_Serial_Numer() ; Takes WMI object as input, returns Local System's Serial Number

EndFunc

Func GET_HDD_Total_and_Free() ; Takes WMI object as input, returns Local System's Hard Drive Total space and Free space

EndFunc

Func GET_Total_RAM() ; Takes WMI object as input, returns Local System's total RAM

EndFunc

Func GET_System_Architecture() ; Takes WMI object as input, returns Local System's Architecture (32 or 64 bit)

EndFunc

Func GET_OS_and_Service_Pack() ; Takes WMI object as input, returns Local System's OS version and Service pack, if any

EndFunc

Func GET_Ethernet_and_Wireless() ; Takes WMI object as input, returns Local System's Wired and Wireless device description and MAC Address

EndFunc

Func GET_Workgroup() ; Takes WMI object as input, returns Local System's Workgroup

EndFunc

