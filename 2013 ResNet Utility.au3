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
#include <array.au3> ; This is provided with Autoit and allows the use of arrays.
#include <Functions\CMDFunc.au3> ;One of my files for common used funtions using command line
#include <Functions\SMARTFunc.au3> ; A function for getting smart data, not my own code
#include <Functions\SystemInfoFunc.au3> ;My function for getting system info
#RequireAdmin ;runs this program as admin, and anything it calls, has admin as well, needed for command prompt scripts
$CurrentVersion = "0.3.0" ; Current version of the software

; Creates GUI, sets name in title bar and icon.
GUICreate("ResNet Utility " & $CurrentVersion, 710, 235) ;Created the GUI form and the size
GUISetIcon("resnet.ico", 0) ;Sets the icon for the window title bar (Should be in the same directory as this file, with this name!)

; Creates Tabs
GUICtrlCreateTab(5, 5, 700, 190) ; Creates tab group
$Tab1 = GUICtrlCreateTabItem("Info") ;Creating the info tab
   ;Creates labels, text boxes and checkboxes to collect system
   ;information and present it.

   ;Display first two columns (OS, System Type, Service Pack)
   GUICtrlCreateLabel("Operating System",22,30) ;Heading
   GUICtrlCreateLabel("OS Edition:",12,52) ;12 is pxels from left, 52, & 
   GUICtrlCreateLabel("System Type:",12,76) ;76,& ^
   GUICtrlCreateLabel("Service Pack:",12,100) ;100 are the height^

   ;Creates labels to contain data
   $OSEd = GUICtrlCreateLabel("xxxx",83,52)
   $SyTy = GUICtrlCreateLabel("xxxx",83,76)
   $SePa = GUICtrlCreateLabel("xxxx",83,100)

   ;Labels for network configuration
   GUICtrlCreateLabel("Network Hardware",165,30) ;Heading
   GUICtrlCreateLabel("Wired Brand:",130,52)
   GUICtrlCreateLabel("Wired MAC:",130,76)
   GUICtrlCreateLabel("Wi-Fi Brand:",130,100)
   GUICtrlCreateLabel("Wi-Fi MAC:",130,124)
   ;GUICtrlCreateLabel("DM Problems",130,178) ;Add Later, not currently implemented

   ;Creates labels to contain data
   $WrBr = GUICtrlCreateLabel("xxxx",200,52,100,12)
   $WrMa = GUICtrlCreateLabel("xxxx",200,76,100)
   $WiBr = GUICtrlCreateLabel("xxxx",200,100,100,12)
   $WiMa = GUICtrlCreateLabel("xxxx",200,124,100)
   ;$DMPr = GUICtrlCreateLabel("xxxx",180,178) ;Add Later, not currently implemented

   ;Labels for Vendor informatrion and System Specs
   GUICtrlCreateLabel("Vendor Information",300,30) ;Heading
   GUICtrlCreateLabel("Brand:",300,52)
   GUICtrlCreateLabel("Serial#:",300,76)
   GUICtrlCreateLabel("HDD:",300,100)
   GUICtrlCreateLabel("RAM:",300,124)

   ;Creates labels to contain data
   $Bran = GUICtrlCreateLabel("xxxx",340,52,100)
   $Seri = GUICtrlCreateLabel("xxxx",340,76,100)
   $HsHf = GUICtrlCreateLabel("xxxx",340,100,100,12)
   $Memo = GUICtrlCreateLabel("xxxx",340,124,100)

   ;Label for Model
   GUICtrlCreateLabel("Model:",450,52)

   ;Data
   $Mode = GUICtrlCreateLabel("xxxx",490,52,100,12)


   ;Group for alerts
   GUICtrlCreateGroup("Alerts",418,107,280,70)
      GUICtrlSetBkColor(-1, 0xFF0000)
	  $Workgroup = GUICtrlCreateLabel("",424,130,100,12) ;Shows up only if workgroup is not ResNet
	  GUICtrlSetColor(-1,0xFF0000)
	  ;$IPAd = GUICtrlCreateLabel("",424,124,100,12) ;Will warn if IP address does not match filter [To be implemented later]
	  ;GUICtrlSetColor(-1,0xFF0000) ; Uncomment this when IPaddress filtering works
   GUICtrlCreateGroup("",-99,-99,1,1)

   ;Sets color of text for data portions
   GUICtrlSetColor($OSEd,0x0000FF)
   GUICtrlSetColor($SyTy,0x0000FF)
   GUICtrlSetColor($SePa,0x0000FF)
   GUICtrlSetColor($WrBr,0x0000FF)
   GUICtrlSetColor($WrMa,0x0000FF)
   GUICtrlSetColor($WiBr,0x0000FF)
   GUICtrlSetColor($WiMa,0x0000FF)
   GUICtrlSetColor($Bran,0x0000FF)
   GUICtrlSetColor($Seri,0x0000FF)
   GUICtrlSetColor($HsHf,0x0000FF)
   GUICtrlSetColor($Memo,0x0000FF)
   GUICtrlSetColor($Mode,0x0000FF)


GUICtrlCreateTabItem("Repair")
   ;Create Buttons to call repair functions (in the repair tab) The On-Click part is handled by the Case statement lower down
   $ResNetwork = GUICtrlCreateButton("Reset Network",87,58,130) ; button that will perform a variety of network connection repair commands
   $RepFirewall = GUICtrlCreateButton("Repair Firewall",87,90,130) ; button to repair the windows firewall
   $ResFirewall = GUICtrlCreateButton("Reset Firewall",87,122,130) ; button to reset the windows firewall settings
   $RepWinUpdate = GUICtrlCreateButton("Repair Windows Update",220,58,130) ; button to run a variety of commands to repair windows update
   $RepPermissions = GUICtrlCreateButton("Repair Permissions",220,90,130) ; button to repair permissions on windows computers
   $FileAssociations = GUICtrlCreateButton("Fix File Associations",220,122,130) ; Runs various commands to repair file associations in the registry
   $SMARTData = GUICtrlCreateButton("Hard Drive Test",353,58,130) ; button to display SMART data for HDD troubleshooting
GUICtrlCreateTabItem("Known Fixes") ; Nothing currently under this tab [to be implemented later]
GUICtrlCreateTabItem("") ; A blank tab item indicates the end of the tab group

; Creates button row at bottom of window
$btnExit = GUICtrlCreateButton("Exit", 5, 200, 60, 30) ; need to remove this button and shift the other three to the left appropriately
$btnMAC = GUICtrlCreateButton("Manual MAC Address",70,200,120,30)
$btnWorkgroup = GUICtrlCreateButton("Change Workgroup",195,200,110,30)
$btnAddRemovePrograms = GUICtrlCreateButton("Programs and Features",310,200,130,30)
; There is plenty of room to add other functionality here, would be a good thing to find out what can be added for automation

; Show GUI
GUISetState(@SW_SHOW) ; Command to actually display the GUI
systemInfo() ;Populates fields on GUI with data obtained from the SystemInfoFunc.au3 file. The function directly makes changes to the values
             ;of the lables on the form, and this needs to be changed [eventually] so that it does not directly change the values from the function
			 ;but instead returns values that are used to update the GUI. That way the function does not need to know the names of variables in this 
			 ;program. Additionally, it probably needs to be several functions instead of just one, for the sake of modularity.

; Up until this point we have seen the interface creation.  None of this previous code has any functions (except systemInfo())
; nor do they call the methods without this next part. 

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
		Case $ResNetwork ;if clicked run vvv
			ResetNetwork() ;CMDFunc Line 10
		Case $RepFirewall ;if clicked...
			RepairFirewall() ;CMDFunc Line 30
		Case $ResFirewall ;if ...
			ResetFirewall() ;CMDFunc Line 36
		Case $RepWinUpdate ;same
			RepairWinUpdate() ;CMDFunc Line 42
		Case $FileAssociations ;same
			FixFileAssociations() ;CMDFunc Line 114
		Case $btnAddRemovePrograms ;same (i'm saying that in all of these seperate cases when the button is clicked it will run the proceeding function)
			AddRemovePrograms() ;CMDFunc Line 26
		Case $SMARTData 
			Initialize_SMART()
	EndSwitch
WEnd ;Case SMARTData is the last case in this while loop, Also it is the only one that uses the other Class(Func) SMARTFunc, so if we can figure out
; if we really need SMARTFunc then we can have this entire 168 lines of code complete.