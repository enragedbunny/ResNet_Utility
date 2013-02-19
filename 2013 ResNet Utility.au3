; Program: ResNet Tool
; Version: 0.3.1
; Author: Johnny Keeton
; Other Contributors: Alex Burgy-Vanhoose, CJ Highley, Josh Back
; Contact: johnny.keeton@gmail.com
; Date Edited: 2013.02.18
; Notes:	1. When placing items, the order is (px from left, px from top, length, height), all sections of this not required
;			2. This program uses external functions, namely CMDFunc.au3, SMARTFunc.au3, SystemInfoFunc.au3 (these are the other files under the
;              "Functions" folder and they must keep the same file name and directory or the program will break (it relies on these files)
;           3. When adding variables, it must start with $ symbol and use capital letters for the first letter of each word in the name
;              Examples:    $ThisIsAVariable $MACAddressVariable, etc
#include <EditConstants.au3> ;For the Tech Notes Box
#include <GuiConstantsEx.au3> ; This file is provided after installing Autoit, and provides the ability to use GUI environment.
#include <WindowsConstants.au3> ; Import variables for tweaking UI. Used to remove borders on GUI and potentially other things.
#include <SendMessage.au3> ;Import ability to send a window a command (Used to move GUI with no borders)
#include <ComboConstants.au3> ;For Combo Boxes
#include <GuiTab.au3> ;For tab colors
#include <Functions\CMDFunc.au3> ;Performs several functions from command line.
#include <Functions\SMARTFunc.au3> ; Gets smart data
#include <Functions\SystemInfoFunc.au3> ;Gets system info
#include <Functions\PreferencesWindow.au3> ;Displays the preferences window
#include <Functions\AboutHelp.au3> ;Displays about and help windows.
#include <Functions\Tab_BK_Color.au3> ;Helps change the background color for tabs
#RequireAdmin ;runs this program as admin, and anything it calls, has admin as well, needed for command prompt scripts

local $ProgramTitle = "ResNet Utility"
local $Version = "0.3.1" ; Current version of the software
local $ReleaseDate = "2012.11.2"
local $HelpFile = "README.txt"
local $lblArray[12] ;Used to set color for data labels in a loop.
Global Const $SC_DRAGMOVE = 0xF012 ;Used for moving the GUI with no borders

; Creates GUI, sets name in title bar and icon.
local $rGUI = GUICreate("ResNet Utility " & $Version, 710, 555,((@DesktopWidth - 710)/2),((@DesktopHeight - 255)/2),$WS_POPUP) ;Created the GUI form and the size. Sets position to center of screen.
GUISetIcon("ResNet.ico", 0) ;Sets the icon for the window title bar (Should be in the same directory as this file, with this name!)

local $objWMI = ObjGet("winmgmts:\\localhost\root\CIMV2") ;Create connection to WMI

;The following rows parse data into arrays for easier use.
local $OSInformation = stringsplit(GET_OS_and_Service_Pack($objWMI),"|") ;Values are as follows (Operating System, Service Pack)
local $BrandModel = stringsplit(GET_Manufacturer_and_Model($objWMI),"|") ;Values are as follows (Computer Manufacturer, Model Number)
local $NetworkSettings = stringSplit(GET_Ethernet_and_Wireless($objWMI),"|") ;The values as follows (Wifi Description, Wifi MAC Address, Wired Description, Wired MAC Address)

; Creates Menu Bar, commented out until more is complete
$mnuFileMenu     = GUICtrlCreateMenu("&File") ;File menu
;$mnuOpenTicket   = GUICtrlCreateMenuItem("Open",$mnuFileMenu) ;Open a ticket from a selected file
;$mnuSaveTicket   = GUICtrlCreateMenuItem("Save",$mnuFileMenu) ;Save current form as a ticket
$mnuExitProgram  = GUICtrlCreateMenuItem("E&xit",$mnuFileMenu) ;Exit the software (Will not autosave ticket unless pref is set to do so (default to save))

$mnuViewMenu     = GUICtrlCreateMenu("&View") ;View menu
;$mnuChecklist    = GUICtrlCreateMenuItem("Checklist Pane",$mnuViewMenu) ;Opens checklist pane
;$mnuTechNotes    = GUICtrlCreateMenuItem("External Tech Notes Pane",$mnuViewMenu) ;Moves Tech Notes to external window
;$mnuTroubleshoot = GUICtrlCreateMenuItem("Troubleshooting Pane",$mnuViewMenu) ;Opens troubleshooting pane

$mnuToolsMenu    = GUICtrlCreateMenu("&Tools") ;Tools menu
$mnuPreferences  = GUICtrlCreateMenuItem("&Preferences",$mnuToolsMenu) ;Opens the preferences window
;$mnuRestart      = GUICtrlCreateMenuItem("Save and Restart",$mnuToolsMenu) ;Saves form and Restart PC

$mnuHelpMenu     = GUICtrlCreateMenu("Help") ;Help Menu  //switched the "?" to "Help" testing git hub tracker and just getting started / ABV 
$mnuAbout        = GUICtrlCreateMenuItem("About",$mnuHelpMenu) ;Opens about window showing version information
$mnuHelp         = GUICtrlCreateMenuItem("Help",$mnuHelpMenu) ;Opens help file for assistance using the program

;Creating Tech Notes Area
$Edit1 = GUICtrlCreateEdit("", 6, 210, 695, 275)
GUICtrlSetData(-1, "Tech Notes")

; Creates Tabs
$hTab_1 = GUICtrlCreateTab(5, 5, 700, 190) ; Creates tab group
GUICtrlCreateTabItem("Info") ;Creating the info tab
	;Creates labels, to collect system information and present it.

	;Display first two columns (OS, System Type, Service Pack)
	GUICtrlCreateLabel("Operating System",22,30) ;Heading
	GUICtrlCreateLabel("OS Edition:",12,52) ;12 is pxels from left, 52, & 
	GUICtrlCreateLabel("System Type:",12,76) ;76,& ^
	GUICtrlCreateLabel("Service Pack:",12,100) ;100 are the height^

	;Creates labels to contain data
	$lblArray[0] = GUICtrlCreateLabel($OSInformation[1],83,52) ;OS Edition (Windows 8, Windows 7, Windows Vista)
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
		GUICtrlCreateLabel("",424,130,100,12) ;Shows up only if workgroup is not ResNet; Not working, need to call function and parse results
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
;GUICtrlCreateTabItem("Known Fixes") ; Nothing currently under this tab [to be implemented later]

$hTab_10 = GUICtrlCreateTabItem("Contact Info") ; Creating Contact Info Tab
	_GUICtrlTab_SetBkColor($rGUI, $hTab_1, 0xFF3333) ;Changeing the background Color
	;Display first column (LN, Hall, EKU Eail)
	GUICtrlCreateLabel("Last Name:",12,50) ;Last Name
	GUICtrlCreateInput("", 85, 48, 130, 20) ;Last Name Input Box
	GUICtrlCreateLabel("Hall:",12,85) ;Hall
	GUICtrlCreateCombo("No Choice", 85, 83, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Hall Combo Box
	GUICtrlSetData(-1, "Burnam|Case|Clay|Combs|Commonwealth|Greek Towers|Keene|Martin|McGregor|Palmer|Sullivan|Telford|Walters|Brockton|Off Campus") ; add other item snd set a new default
	GUICtrlCreateLabel("EKU Email:",12,120) 
	GUICtrlCreateInput("", 85, 118, 130, 20) ;EMail Input Box
	
	;Display second column (FN, Phone)
	GUICtrlCreateLabel("First Name:",222,50) ;First Name
	GUICtrlCreateInput("", 295, 48, 130, 20) ;F N input box
	GUICtrlCreateLabel("Phone:",222,85) ; Phone Number
	GUICtrlCreateInput("", 295, 83, 130, 20) ;Phone input box
	
	;Display Third column (EKU ID, Phone; Other)
	GUICtrlCreateLabel("EKU ID:",442,50) ;EKU ID
	GUICtrlCreateInput("", 515, 48, 130, 20) ;EKU ID input box
	GUICtrlCreateLabel("Phone Other:",442,85) ; Secondary Phone Number
	GUICtrlCreateInput("", 515, 83, 130, 20) ;2nd Phone input box
	
$hTab_11 = GUICtrlCreateTabItem("Incident Notes") ; Creating Incident Notes Tab
	_GUICtrlTab_SetBkColor($rGUI, $hTab_1, 0x33CC33) ;Changeing the background Color
	;Display first column (LN, Hall, EKU Eail)
	GUICtrlCreateLabel("Problem Type:",12,50) ;Problem Type
	$Problem = GUICtrlCreateCombo("No Choice", 85, 48, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Problem type Combo Box
	GUICtrlSetData(-1, "Blackboard|Cart Checkout|Clusters|Complaint-Esc. to Lisa|Drive Recovery|Ekey|EKU Direct|Electronics Recycling|Email|Employment|Game Console|Hardware-NIC|Hardware-Other|IM|In-Room Pickup|Judicial Sanction|Labs-Combs 230|Marketing, PR|Network-Bypass|Network-Connection|Phone|Pickup In-Room|SkyDrive|Software-Antivirus|Software-Burned|Software-CCA|Software-OS Updates|Software-Other|Software-Virus/Malware|Voicemail|W: Drive|Wiring ")
	GUICtrlCreateLabel("Platform:",12,85) ;Platform
	$Platform = GUICtrlCreateCombo("No Choice", 85, 83, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Platform Combo Box
	GUICtrlSetState($Platform, $GUI_DISABLE) ; Disable the combo until a Symptom is chosen <<<<<<<<<<<<<<<<<<<<
	GUICtrlCreateLabel("Symptom:",222,50) ;Symptom
	$Symptom = GUICtrlCreateCombo("No Choice", 295, 48, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Symptom Combo Box
	GUICtrlSetState($Symptom, $GUI_DISABLE) ; Disable the combo until a Platform is chosen <<<<<<<<<<<<<<<<<<<<
	GUICtrlCreateLabel("Root Cause:",222,85) ; Root Cause
	$RootCause = GUICtrlCreateCombo("No Choice", 295, 83, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Root Cause Combo Box
	GUICtrlSetState($RootCause, $GUI_DISABLE) ; Disable the combo until a Platform is chosen <<<<<<<<<<<<<<<<<<<<
	GUICtrlCreateCheckbox("Kevin's Cell (Wiring)",442,50) ;Kevins Cell check
	
$hTab_12 = GUICtrlCreateTabItem("Triage Notes") ; Creating Triage Notes Tab 
	_GUICtrlTab_SetBkColor($rGUI, $hTab_1, 0xFFCC33) ;Changeing the background Color
	;Display Top Row (Start Tech, & Start up issues)
	GUICtrlCreateLabel("Start Tech: ",12,50) ;Start Tech
	GUICtrlCreateInput("", 85, 48, 130, 20) ;x, y, width, height
	GUICtrlCreateLabel("Startup issues: ",222,50) ;Startup issues 
	GUICtrlCreateInput("", 295, 48, 390, 20) ;Startup issues input box
	GUICtrlCreateLabel("Triage Notes: ",12,85) ;Triage Notes
	GUICtrlCreateEdit("", 15, 100, 680, 85) ;Triage Notes Text Box
	
$hTab_13 = GUICtrlCreateTabItem("Closing Notes") ; Creating Triage Notes Tab
	_GUICtrlTab_SetBkColor($rGUI, $hTab_1, 0x0066FF) ;Changeing the background Color
	GUICtrlCreateLabel("Issues:",12,50) ;Issues
	GUICtrlCreateCombo("No Choice", 85, 48, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Combo Box First Try ABV
	GUICtrlSetData(-1, "Drivers|Judicial Cleanup|Malware|MS Update|NAC Issues|New Connect|Other|Reconnect|Virus|Wireless")
	GUICtrlCreateLabel("Success or",12,85) ;Success or Defect
	GUICtrlCreateLabel("Defect?",12,100)
	GUICtrlCreateCombo("No Choice", 85, 83, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Platform Combo Box
	GUICtrlSetData(-1, "Success|Defect")
	
	GUICtrlCreateLabel("Other Issues?",222,50) ;Other Issues
	GUICtrlCreateInput("", 295, 48, 130, 20) ;Other Issues Input Box
	GUICtrlCreateLabel("Defect Why?",222,85) ;Defect Box 
	GUICtrlCreateInput("", 295, 83, 130, 20) ;Defect input box
	
	GUICtrlCreateLabel("Traige End Tech:",442,50) ;Triage End Tech
	GUICtrlCreateCombo("Default to Prefs", 530, 48, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Triage end tech Combo Box
	GUICtrlCreateLabel("Closing Tech:",442,85) ; Closing Tech
	GUICtrlCreateCombo("Lead Tech (drop off)", 530, 83, 130, 20, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL)) ;Closing tech Combo Box
	
GUICtrlCreateTabItem("") ; A blank tab item indicates the end of the tab group

; Creates button row at bottom of window
local $btnMAC = GUICtrlCreateButton("Manual MAC Address",5,500,120,30) ;Button to open cmd prompt, type ipconfig/all, and let you manually check network settings
local $btnWorkgroup = GUICtrlCreateButton("Change Workgroup",130,500,110,30) ;Button to open advanced computer settings, clicks change so you can manually change workgroup settings
local $btnAddRemovePrograms = GUICtrlCreateButton("Programs and Features",245,500,130,30) ;Opens Programs and Features to manually uninstall programs
local $btnSave = GUICtrlCreateButton("Save",515,500,50,30) ;Button to eventually save progress No action in While loop blow as well as no function in CMDFunc
local $btnExport = GUICtrlCreateButton("Export",570,500,65,30) ;Button for Exporting No action or fucntion yet
local $btnUpload = GUICtrlCreateButton("Upload",640,500,65,30) ;Button for Uploading Ticket No action or function yet

GUISetState(@SW_SHOW) ;Command to actually display the GUI

While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			Exit
		Case $GUI_EVENT_CLOSE ;Closes window if program is given close signal
			Exit ;This Exit command is what actually makes the program exit.
		Case $mnuExitProgram
			Exit ;This Exit command is what actually makes the program exit.
		Case $GUI_EVENT_PRIMARYDOWN
			_SendMessage($rGUI, $WM_SYSCOMMAND, $SC_DRAGMOVE, 0)
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
		;Case $mnuOpenTicket ;if this menu item is clicked
			;OpenTicket() ;Loads ticket into window
		;Case $mnuSaveTicket ;if this menu item is clicked
			;SaveTicket() ;Saves current form into ticket file
		Case $mnuPreferences ;if this menu item is clicked
			CreatePreferencesWindow() ;Creates GUI to set defaults and window preferences
		;Case $mnuRestart ;if this menu item is clicked
			;SaveTicket() ;Saves current form into ticket file
			;RestartPC() ;Restarts PC
		;Case $mnuChecklist ;if this menu item is clicked
			;CreateChecklistWindow() ;Creates GUI checklist for walk-in/drop-off procedure
		;Case $mnuTechNotes ;if this menu item is clicked
			;CreateTechNotesWindow() ; Move tech notes to external window
		;Case $mnuTroubleshoot ;if this menu item is clicked
			;CreateTroubleshootWindow() ;Creates GUI for network troubleshooting
		Case $mnuAbout ;if this menu item is clicked
			CreateAboutWindow($ProgramTitle,$Version,$ReleaseDate) ;Displays program information
		Case $mnuHelp ;if this menu item is clicked
			CreateHelpWindow($HelpFile) ;Displays information about how to use software
		Case $Problem
            GUICtrlSetState($Symptom, $GUI_ENABLE) ; Enable the Symptom combo box<<<<<<<<<<<<<<<<<<
		Case $Symptom
            GUICtrlSetState($Platform, $GUI_ENABLE) ; Enable the Platform combo box<<<<<<<<<<<<<<<<<<
		Case $Platform
            GUICtrlSetState($RootCause, $GUI_ENABLE) ; Enable the Platform combo box<<<<<<<<<<<<<<<<<<
	EndSwitch
WEnd