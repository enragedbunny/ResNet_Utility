; 1st Program I think is easiest to figure out what is going on, and get an idea of some Autoit script, and have the ResNet tool running to learn.
; Program: ResNet Tool
; Version: 0.3.0
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com [Please everyone add their contact email in this section]
; Date Edited: 2012.10.06
; Purpose:
; Notes:	1. When placing items, the order is (px from left, px from top, length, height), all sections of this not required
;			2. This program uses external functions, namely CMDFunc.au3, SMARTFunc.au3, SystemInfoFunc.au3 (these are the other files under the
;              "Functions" folder and they must keep the same file name and directory or the program will break (it relies on these files)
;           3. When adding variables, it must start with $ symbol and use capital letters for the first letter of each word in the name
;              Examples:    $ThisIsAVariable $MACAddressVariable, etc
;           4. Although not perfectly implemented yet, the goal is to completely separate GUI from implementation, and localize variables as much
;              as possible
#include <GuiConstantsEx.au3> ; This file is provided after installing Autoit, and provides the ability to use GUI environment.
#include <Functions\CMDFunc.au3> ;One of my files for common used funtions using command line
#include <Functions\SMARTFunc.au3> ; A function for getting smart data, not my own code
#include <Functions\SystemInfoFunc.au3> ;My function for getting system info
#RequireAdmin ;runs this program as admin, and anything it calls, has admin as well, needed for command prompt scripts

local $CurrentVersion = "0.3.0" ; Current version of the software
local $Tab1,$Tab2,$Tab3 ; Declare Tab item variables
local $OSEd,$SyTy,$SePa,$WaBr,$WaMa,$WiBr,$WiMa,$Bran,$Seri,$HsHf,$Memo,$Mode ; Declare variables to store system information
local $Workgroup ; Declare variable to store workgroup in
local $lblArray[12] ;Used to set color for data labels in a loop.

; Creates GUI, sets name in title bar and icon.
GUICreate("ResNet Utility " & $CurrentVersion, 710, 235) ;Created the GUI form and the size
GUISetIcon("resnet.ico", 0) ;Sets the icon for the window title bar (Should be in the same directory as this file, with this name!)
local $objWMI = ObjGet("winmgmts:\\localhost\root\CIMV2") ;Create connection to WMI

;The following rows parse data into arrays for easier use.
local $OSInformation = stringsplit(GET_OS_and_Service_Pack($objWMI),"|") ;Values are as follows (Operating System, Service Pack)
local $BrandModel = stringsplit(GET_Manufacturer_and_Model($objWMI),"|") ;Values are as follows (Computer Manufacturer, Model Number)
local $NetworkSettings = stringSplit(GET_Ethernet_and_Wireless($objWMI),"|") ;The values as follows (Wifi Description, Wifi MAC Address, Wired Description, Wired MAC Address)

; Creates Tabs
GUICtrlCreateTab(5, 5, 700, 190) ; Creates tab group
$Tab1 = GUICtrlCreateTabItem("Info") ;Creating the info tab
	;Creates labels, to collect system information and present it.

	;Display first two columns (OS, System Type, Service Pack)
	GUICtrlCreateLabel("Operating System",22,30) ;Heading
	GUICtrlCreateLabel("OS Edition:",12,52) ;12 is pxels from left, 52, & 
	GUICtrlCreateLabel("System Type:",12,76) ;76,& ^
	GUICtrlCreateLabel("Service Pack:",12,100) ;100 are the height^

	;Creates labels to contain data
	$lblArray[0] = GUICtrlCreateLabel($OSInformation[1],83,52) ;OS Edition (Windows 8, Windows 7, Windows Vista
	$lblArray[1] = GUICtrlCreateLabel(GET_System_Architecture($objWMI),83,76) ;32 or 64-bit
	$lblArray[2] = GUICtrlCreateLabel($OSInformation[2],83,100) ;Service pack version installed

	;Labels for network configuration
	GUICtrlCreateLabel("Network Hardware",165,30) ;Heading
	GUICtrlCreateLabel("Wired Brand:",130,52)
	GUICtrlCreateLabel("Wired MAC:",130,76)
	GUICtrlCreateLabel("Wi-Fi Brand:",130,100)
	GUICtrlCreateLabel("Wi-Fi MAC:",130,124)
	;GUICtrlCreateLabel("DM Problems",130,178) ;Add Later, not currently implemented

	;Creates labels to contain data
	$lblArray[3] = GUICtrlCreateLabel($NetworkSettings[3],200,52,100,12) ;Wired Brand and description
	$lblArray[4] = GUICtrlCreateLabel($NetworkSettings[4],200,76,100) ;Wired MAC Address
	$lblArray[5] = GUICtrlCreateLabel($NetworkSettings[1],200,100,100,12) ;Wireless Brand and description
	$lblArray[6] = GUICtrlCreateLabel($NetworkSettings[2],200,124,100) ;Wireless MAC Address
	;GUICtrlCreateLabel("xxxx",180,178) ;Add Later, not currently implemented

	;Labels for Vendor informatrion and System Specs
	GUICtrlCreateLabel("Vendor Information",300,30) ;Heading
	GUICtrlCreateLabel("Brand:",300,52)
	GUICtrlCreateLabel("Serial#:",300,76)
	GUICtrlCreateLabel("HDD:",300,100)
	GUICtrlCreateLabel("RAM:",300,124)

	;Creates labels to contain data
	$lblArray[7] = GUICtrlCreateLabel($BrandModel[1],340,52,100) ;Computer Manufacturer
	$lblArray[8] = GUICtrlCreateLabel(GET_Serial_Number($objWMI),340,76,100) ;Serial Number
	$lblArray[9] = GUICtrlCreateLabel(GET_HDD_Total_and_Free($objWMI),340,100,100,12) ;Hard drive total and free space
	$lblArray[10] = GUICtrlCreateLabel(GET_Total_RAM($objWMI),340,124,100) ;Total RAM on system

	;Label for Model
	GUICtrlCreateLabel("Model:",450,52)

	;Data
	$lblArray[11] = GUICtrlCreateLabel($BrandModel[2],490,52,100,12) ;Model Number


	;Group for alerts
	GUICtrlCreateGroup("Alerts",418,107,280,70)
		GUICtrlSetBkColor(-1, 0xFF0000)
		GUICtrlCreateLabel("",424,130,100,12) ;Shows up only if workgroup is not ResNet
		GUICtrlSetColor(-1,0xFF0000)
		;$IPAd = GUICtrlCreateLabel("",424,124,100,12) ;Will warn if IP address does not match filter [To be implemented later]
		;GUICtrlSetColor(-1,0xFF0000) ; Uncomment this when IPaddress filtering works
	GUICtrlCreateGroup("",-99,-99,1,1)
   
	local $i ;declares variable for loop
	If IsArray($lblArray) Then
		For $i = 0 to 11
			GUICtrlSetColor($lblArray[$i],0x0000FF)
		Next
	EndIf

$Tab2 = GUICtrlCreateTabItem("Repair")
	;Create Buttons to call repair functions (in the repair tab) The On-Click part is handled by the Case statement lower down
	local $btnResNetwork = GUICtrlCreateButton("Reset Network",87,58,130) ; button that will perform a variety of network connection repair commands
	local $btnRepFirewall = GUICtrlCreateButton("Repair Firewall",87,90,130) ; button to repair the windows firewall
	local $btnResFirewall = GUICtrlCreateButton("Reset Firewall",87,122,130) ; button to reset the windows firewall settings
	local $btnRepWinUpdate = GUICtrlCreateButton("Repair Windows Update",220,58,130) ; button to run a variety of commands to repair windows update
	local $btnRepPermissions = GUICtrlCreateButton("Repair Permissions",220,90,130) ; button to repair permissions on windows computers
	local $btnFileAssociations = GUICtrlCreateButton("Fix File Associations",220,122,130) ; Runs various commands to repair file associations in the registry
	local $btnSMARTData = GUICtrlCreateButton("Hard Drive Test",353,58,130) ; button to display SMART data for HDD troubleshooting
$Tab3 = GUICtrlCreateTabItem("Known Fixes") ; Nothing currently under this tab [to be implemented later]
GUICtrlCreateTabItem("") ; A blank tab item indicates the end of the tab group

; Creates button row at bottom of window
local $btnExit = GUICtrlCreateButton("Exit", 5, 200, 60, 30) ; need to remove this button and shift the other three to the left appropriately
local $btnMAC = GUICtrlCreateButton("Manual MAC Address",70,200,120,30)
local $btnWorkgroup = GUICtrlCreateButton("Change Workgroup",195,200,110,30)
local $btnAddRemovePrograms = GUICtrlCreateButton("Programs and Features",310,200,130,30)
; There is plenty of room to add other functionality here, would be a good thing to find out what can be added for automation

GUISetState(@SW_SHOW) ;Command to actually display the GUI

; Up until this point we have seen the interface creation.  None of this previous code has any functions nor do they call the methods without this next part. 
While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE ;Closes window if program is given close signal (x at top of screen), built it 
			Exit ;This Exit command is what actually makes the program exit.
		Case $btnExit ; Same, just for our custom button (probably don't need). Remove this once the Exit button has been removed.
			Exit
		Case $btnWorkgroup ;if this button is clicked on then it will run vvvvvvv
			OpenWorkgroup() ;This method is located in the Function: CMDFunc on Line 131
		Case $btnMAC ;if clicked run vvv
			AltGetMAC() ; Alternate way to get MAC address, CMDFunc Line 6
		Case $btnResNetwork ;if clicked run vvv
			ResetNetwork() ;CMDFunc Line 10
		Case $btnRepFirewall ;if clicked...
			RepairFirewall() ;CMDFunc Line 30
		Case $btnResFirewall ;if ...
			ResetFirewall() ;CMDFunc Line 36
		Case $btnRepWinUpdate ;same
			RepairWinUpdate() ;CMDFunc Line 42
		Case $btnFileAssociations ;same
			FixFileAssociations() ;CMDFunc Line 114
		Case $btnAddRemovePrograms ;same (i'm saying that in all of these seperate cases when the button is clicked it will run the proceeding function)
			AddRemovePrograms() ;CMDFunc Line 26
		Case $btnSMARTData 
			Initialize_SMART()
	EndSwitch
WEnd ;Case SMARTData is the last case in this while loop, Also it is the only one that uses the other Class(Func) SMARTFunc, so if we can figure out
; if we really need SMARTFunc then we can have this entire 168 lines of code complete.