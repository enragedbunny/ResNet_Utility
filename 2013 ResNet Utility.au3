; Program: ResNet Tool
; Version: 0.3.1
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com [Please everyone add their contact email in this section]
; Date Edited: 2012.10.29
; Notes:	1. When placing items, the order is (px from left, px from top, length, height), all sections of this not required
;			2. This program uses external functions, namely CMDFunc.au3, SMARTFunc.au3, SystemInfoFunc.au3 (these are the other files under the
;              "Functions" folder and they must keep the same file name and directory or the program will break (it relies on these files)
;           3. When adding variables, it must start with $ symbol and use capital letters for the first letter of each word in the name
;              Examples:    $ThisIsAVariable $MACAddressVariable, etc
#include <GuiConstantsEx.au3> ; This file is provided after installing Autoit, and provides the ability to use GUI environment.
#include <Functions\CMDFunc.au3> ;One of my files for common used funtions using command line
#include <Functions\SMARTFunc.au3> ; A function for getting smart data, not my own code
#include <Functions\SystemInfoFunc.au3> ;My function for getting system info
#include <Functions\PreferencesWindow.au3> ;My Function for displaying the preferences window
#RequireAdmin ;runs this program as admin, and anything it calls, has admin as well, needed for command prompt scripts

local $CurrentVersion = "0.3.1" ; Current version of the software
local $lblArray[12] ;Used to set color for data labels in a loop.

; Creates GUI, sets name in title bar and icon.
GUICreate("ResNet Utility " & $CurrentVersion, 710, 255) ;Created the GUI form and the size
GUISetIcon("resnet.ico", 0) ;Sets the icon for the window title bar (Should be in the same directory as this file, with this name!)
local $objWMI = ObjGet("winmgmts:\\localhost\root\CIMV2") ;Create connection to WMI

;The following rows parse data into arrays for easier use.
local $OSInformation = stringsplit(GET_OS_and_Service_Pack($objWMI),"|") ;Values are as follows (Operating System, Service Pack)
local $BrandModel = stringsplit(GET_Manufacturer_and_Model($objWMI),"|") ;Values are as follows (Computer Manufacturer, Model Number)
local $NetworkSettings = stringSplit(GET_Ethernet_and_Wireless($objWMI),"|") ;The values as follows (Wifi Description, Wifi MAC Address, Wired Description, Wired MAC Address)

; Creates Menu Bar 
$mnuFileMenu     = GUICtrlCreateMenu("&File") ;File menu
$mnuOpenTicket   = GUICtrlCreateMenuItem("Open",$mnuFileMenu) ;Open a ticket from a selected file
$mnuSaveTicket   = GUICtrlCreateMenuItem("Save",$mnuFileMenu) ;Save current form as a ticket
$mnuExitProgram  = GUICtrlCreateMenuItem("E&xit",$mnuFileMenu) ;Exit the software (Will not autosave ticket unless pref is set to do so (default to save))

$mnuViewMenu     = GUICtrlCreateMenu("&View") ;View menu
$mnuChecklist    = GUICtrlCreateMenuItem("Checklist Pane",$mnuViewMenu) ;Opens checklist pane
$mnuTechNotes    = GUICtrlCreateMenuItem("External Tech Notes Pane",$mnuViewMenu) ;Moves Tech Notes to external window
$mnuTroubleshoot = GUICtrlCreateMenuItem("Troubleshooting Pane",$mnuViewMenu) ;Opens troubleshooting pane

$mnuToolsMenu    = GUICtrlCreateMenu("&Tools") ;Tools menu
$mnuPreferences  = GUICtrlCreateMenuItem("&Preferences",$mnuToolsMenu) ;Opens the preferences window
$mnuRestart      = GUICtrlCreateMenuItem("Save and Restart",$mnuToolsMenu) ;Saves form and Restart PC

$mnuHelpMenu     = GUICtrlCreateMenu("Help") ;Help Menu
$mnuAbout        = GUICtrlCreateMenuItem("About",$mnuHelpMenu) ;Opens about window showing version information
$mnuHelp         = GUICtrlCreateMenu("Help",$mnuHelpMenu) ;Opens help file for assistance using the program

; Creates Tabs
GUICtrlCreateTab(5, 5, 700, 190) ; Creates tab group
GUICtrlCreateTabItem("Info") ;Creating the info tab
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

	;Labels for Vendor information and System Specs
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
	GUICtrlCreateGroup("Alerts",418,107,280,70)  ;Creates an alert box to display information alerts
		GUICtrlSetBkColor(-1, 0xFF0000) ;Makes the border red
		GUICtrlCreateLabel("",424,130,100,12) ;Shows up only if workgroup is not ResNet
		GUICtrlSetColor(-1,0xFF0000) ;sets the color of the workgroup alert to red
		;$IPAd = GUICtrlCreateLabel("",424,124,100,12) ;Will warn if IP address does not match filter [To be implemented later]
		;GUICtrlSetColor(-1,0xFF0000) ; Uncomment this when IPaddress filtering works
	GUICtrlCreateGroup("",-99,-99,1,1)
   
	local $i ;declares variable for loop
	If IsArray($lblArray) Then ;Makes sure the array was created correctly
		For $i = 0 to 11 ;Loops through the data labels (created as array previously)
			GUICtrlSetColor($lblArray[$i],0x0000FF) ;Sets the color of the text for 'data' in the GUI to blue
		Next
	EndIf

GUICtrlCreateTabItem("Repair")
	;Create Buttons to call repair functions (in the repair tab) The On-Click part is handled by the Case statement lower down
	local $btnResNetwork = GUICtrlCreateButton("Reset Network",87,58,130) ; button that will perform a variety of network connection repair commands
	local $btnRepFirewall = GUICtrlCreateButton("Repair Firewall",87,90,130) ; button to repair the windows firewall
	local $btnResFirewall = GUICtrlCreateButton("Reset Firewall",87,122,130) ; button to reset the windows firewall settings
	local $btnRepWinUpdate = GUICtrlCreateButton("Repair Windows Update",220,58,130) ; button to run a variety of commands to repair windows update
	local $btnRepPermissions = GUICtrlCreateButton("Repair Permissions",220,90,130) ; button to repair permissions on windows computers
	local $btnFileAssociations = GUICtrlCreateButton("Fix File Associations",220,122,130) ; Runs various commands to repair file associations in the registry
	local $btnSMARTData = GUICtrlCreateButton("Hard Drive Test",353,58,130) ; button to display SMART data for HDD troubleshooting
GUICtrlCreateTabItem("Known Fixes") ; Nothing currently under this tab [to be implemented later]
GUICtrlCreateTabItem("") ; A blank tab item indicates the end of the tab group

; Creates button row at bottom of window
local $btnMAC = GUICtrlCreateButton("Manual MAC Address",5,200,120,30) ;Button to open cmd prompt, type ipconfig/all, and let you manually check network settings
local $btnWorkgroup = GUICtrlCreateButton("Change Workgroup",130,200,110,30) ;Button to open advanced computer settings, clicks change so you can manually change workgroup settings
local $btnAddRemovePrograms = GUICtrlCreateButton("Programs and Features",245,200,130,30) ;Opens Programs and Features to manually uninstall programs
; There is plenty of room to add other functionality here, would be a good thing to find out what can be added for automation

GUISetState(@SW_SHOW) ;Command to actually display the GUI

; Up until this point we have seen the interface creation.  None of this previous code has any functions nor do they call the methods without this next part. 
While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE ;Closes window if program is given close signal
			Exit ;This Exit command is what actually makes the program exit.
		Case $btnWorkgroup ;if this button is clicked
			OpenWorkgroup() ;Opens the advanced computer settings and clicks button to change workgroup
		Case $btnMAC ;if this button is clicked
			AltGetMAC() ; Opens cmd prompt and types ipconfig/all
		Case $btnResNetwork ;if this button is clicked
			ResetNetwork() ;Runs multiple commands at cmd prompt to repair network settings
		Case $btnRepFirewall ;if this button is clicked
			RepairFirewall() ;runs commands to repair Windows firewall settings
		Case $btnResFirewall ;if this button is clicked
			ResetFirewall() ;resets any configuration done to Windows firewall
		Case $btnRepWinUpdate ;if this button is clicked
			RepairWinUpdate() ;Runs multiple phases of commands to repair Windows Update
		Case $btnFileAssociations ;if this button is clicked
			FixFileAssociations() ;Runs some commands to fix file associations (.exe, .lnk, etc)
		Case $btnAddRemovePrograms ;if this button is clicked
			AddRemovePrograms() ;Opens Add/Remove programs or in vista/7, Programs and features.
		Case $btnSMARTData ;if this button is clicked
			Initialize_SMART() ;Opens a new window with SMART information for C: drive
		Case $mnuOpenTicket ;if this menu item is clicked
			OpenTicket() ;Loads ticket into window
		Case $mnuSaveTicket ;if this menu item is clicked
			SaveTicket() ;Saves current form into ticket file
		Case $mnuPreferences ;if this menu item is clicked
			CreatePreferencesWindow() ;Creates GUI to set defaults and window preferences
		Case $mnuRestart ;if this menu item is clicked
			SaveTicket() ;Saves current form into ticket file
			RestartPC() ;Restarts PC
		Case $mnuChecklist ;if this menu item is clicked
			CreateChecklistWindow() ;Creates GUI checklist for walk-in/drop-off procedure
		Case $mnuTechNotes ;if this menu item is clicked
			CreateTechNotesWindow() ; Move tech notes to external window
		Case $mnuTroubleshoot ;if this menu item is clicked
			CreateTroubleshootWindow() ;Creates GUI for network troubleshooting
		Case $mnuAbout ;if this menu item is clicked
			CreateAboutWindow() ;Displays program information
		Case $mnuHelp ;if this menu item is clicked
			CreateHelpWindow() ;Displays information about how to use software
	EndSwitch
WEnd