;This file contains many functions for doing various things from command line

Func AltGetMAC() ; Opens cmd.exe and executes ipconfig /all to manually check IP and MAC Addresses
	RunWait('"' & @ComSpec & '" /k ' & 'ipconfig /all', @SystemDir,@SW_SHOW) ; RunWait starts the indicated program and waits until it loads before executing other code.
EndFunc ;Ends Function

Func ResetNetwork() ; Runs a variety of commands to repair/reset network settings to default values
	local $i ;Declares $i variable for loop
	local $Data = 'netsh winsock reset|' & _                          ;Resets Winsock settings
				 'netsh winsock reset catalog|' & _                   ;Resets Winsock catalog
				 'netsh interface ip reset c:\int-resetlog.txt|' & _  ;Reset IP Interface settings and save a log
				 'netsh interface ip delete arpcache|' & _            ;Delete ARP cache on all IP interfaces
				 'ipconfig /flushdns|' & _                            ;Flush DNS
				 'ipconfig /registerdns'                              ;Register DNS
	local $CMD = StringSplit($Data,"|") ;Splits $Data into an array called $CMD, using | symbol as delimeter
	
	If IsArray($CMD) Then ;checks if $CMD is an array then runs the loop.
		For $i = 1 to $CMD[0] ;looping from 1 to the end of the array because CMD[0] is the length from string split
				RunWait('"' & @ComSpec & '" /c ' & $CMD[$i], @SystemDir,@SW_HIDE) ;@ComSpec is cmd prompt, and /c is for closing it. @SW_HIDE is a autoit command that hides the window, starts it hidden. @SystemDir is the working directory (not the path to the file).
		Next ;Next continues the loop
	EndIf ;EndIf ends this If statement
EndFunc ;Ends Function

Func AddRemovePrograms() ; Opens the Programs and Features window
	RunWait('"' & @ComSpec & '" /c ' & 'appwiz.cpl', @SystemDir,@SW_HIDE) ;types appwiz.cpl in cmd prompt: runs add remove programs
EndFunc ;Ends Function

Func RepairFirewall() ; Runs a command to repair the windows firewall settings
   local $CMD1 = 'rundll32.exe setupapi.dll,InstallHinfSection Ndi-Steelhead 132 %windir%\inf\netrass.inf' ;Repairs windows firewall
   RunWait('"' & @ComSpec & '" /c ' & $CMD1, @SystemDir,@SW_HIDE) ;Runs the previous command from command prompt
EndFunc ;Ends Function

Func ResetFirewall() ; Flushes out windows firewall settings (sets to default settings)
   local $CMD1 = 'netsh firewall reset' ;Runs this command in command prompt

   RunWait('"' & @ComSpec & '" /c ' & $CMD1, @SystemDir,@SW_HIDE)
EndFunc ;Ends Function

Func RepairWinUpdate() ; Runs three sets of commands to repair windows update
	Local $i ;Declares $i variable for loop
	;The first data set stops some WU services, deletes some files, renames some files/folders and changes the directory path to System32 folder
	Local $Data1 = 'net stop bits|' & _ ;& _ continues the next line... eg: net stop bits|net stop wuauserv|...
			     'net stop wuauserv|' & _   ;Stops wuauserv windows service
			     'Del "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat"|' & _ ;Delete qmgr*.dat files
				 'Ren %systemroot%\SoftwareDistrobution\DataStore *.bak|' & _ ;Rename DataStore files/Folders to *.bak
				 'Ren %systemroot%\SoftwareDistrobution\Download *.bak|' & _ ;Rename Download Folder to Download.bak
				 'Ren %systemroot%\system32\catroot2 *.bak|' & _ ;Rename any files/folders that begin with catroot and change to *.bak
				 'sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|' & _  ;Configure BITS settings (a windows service)
				 'sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|' & _ ;Configure WUAUServ settings (a windows service)
				 'cd /d %WINDIR%\system32' ;Changes working directory in cmd prompt to run the following DLL Register commands
	;The second data set reregisters some dll files (if present) in the system32 folder (These DLLs are required for windows update to function)
	Local $Data2 = 'atl.dll|' & _
				 'urlmon.dll|' & _
				 'mshtml.dll|' & _
				 'shdocvw.dll|' & _
				 'browseui.dll|' & _
				 'jscript.dll|' & _  ;DLL file associated with javascript
				 'vbscript.dll|' & _ ;DLL file associated with vbscript
				 'scrrun.dll|' & _
				 'msxml.dll|' & _
				 'msxml3.dll|' & _
				 'msxml6.dll|' & _
				 'actxprxy.dll|' & _
				 'softpub.dll|' & _
				 'wintrust.dll|' & _
				 'dssenh.dll|' & _
				 'rsaenh.dll|' & _
				 'gpkcsp.dll|' & _
				 'sccbase.dll|' & _
				 'slbcsp.dll|' & _
				 'cryptdlg.dll|' & _
				 'oleaut32.dll|' & _
				 'ole32.dll|' & _
				 'shell32.dll|' & _
				 'initpki.dll|' & _
				 'wuapi.dll|' & _
				 'wuaueng.dll|' & _
				 'wuaueng1.dll|' & _
				 'wucltui.dll|' & _
				 'wups.dll|' & _
				 'wups2.dll|' & _
				 'wuweb.dll|' & _
				 'qmgr.dll|' & _
				 'qmgrprxy.dll|' & _
				 'wucltux.dll|' & _
				 'muweb.dll|' & _
				 'wuwebv.dll'
	;The third (and final) data set does some cleanup by resetting winsock settings and starting the windows update services back up.
	Local $Data3 = 'netsh reset winsock|' & _ ;Reset Winsock settings
				  'net start bits|' & _ ;Start BITS windows service
				  'net start wuauserv' ;Start wuauserv windows service
	Local $CMD = StringSplit($Data1, "|") ;Converts $Data1 into array $CMD
	Local $CMDReg = StringSplit($Data2, "|") ; Converts $Data2 into array $CMDReg
	Local $CMD2 = StringSplit($Data3, "|") ; Converts $Data3 into array $CMD2
   
	If IsArray($CMD) Then ;checks if $CMD is an array then runs the loop.
		For $i = 1 to $CMD[0] ;looping from 1 to the end of the array because CMD[0] is the length from string split
			RunWait('"' & @ComSpec & '" /c ' & $CMD[$i], @SystemDir, @SW_HIDE) ;This loopp is running all of the functions in Data1.
		Next ;Next continues the loop
	EndIf ;EndIf ends this if statement
	;so once this loop is closed it will run this next data set vvvvv

	If IsArray($CMDReg) Then ;Checks if $CMDReg was properly created as array
		For $i = 1 to $CMDReg[0] ;Loops from 1 to end of array ($CMDReg[0] contains the length of the array)
			RunWait('"' & @ComSpec & '" /c ' & 'regsvr32.exe /s ' & $CMDReg[$i], @SystemDir, @SW_HIDE)
		Next ;Next continues the loop
	EndIf ;EndIf ends this if statement

	If IsArray($CMD2) Then ;Checks if $CMD2 was properly created as array
		For $i = 1 to $CMD2[0] ;Loops from 1 to end of array ($CMDReg[0] contains the length of the array)
			RunWait('"' & @ComSpec & '" /c ' & $CMD2[$i], @SystemDir, @SW_HIDE)
		Next ;Next continues the loop
	EndIf ;EndIf ends this if statement
EndFunc ;Ends Function

Func FixFileAssociations() ;Fixes file extention problems (needs to be expanded and improved a ton)
	Local $i ;Declares $i variable for loop
	;The following dataset is used to set file extentions back to default (alters the registry)
	Local $Data = '.exe=exefile|' & _  ;creates an association between .exe extensions and the program that opens exe Files
			     '.bat=batfile|' & _   ;creates an association between .bat extensions and the program that opens bat Files
				 '.reg=regfile|' & _   ;creates an association between .reg extensiosn and the program that opens reg Files
				 '.com=comfile|' & _   ;creates an association between .com extensions and the program that opens com Files
				 '.xml=xmlfile|' & _   ;creates an association between .xml extensions and the program that opens xml Files
				 '.lnk=lnkfile|' & _   ;creates an association between .lnk (shortcuts) extensions and the program that opens lnk Files
				 '.ico=icofile|'       ;creates an association between .ico (icons) extensions and the program that opens ico Files
	Local $CMD = StringSplit($Data, "|") ;String split saves length of array as CMD[0] in the array.
	If IsArray($CMD) Then ;Checks if $CMD was properly created as array
		For $i = 1 to $CMD[0] ;Loops from 1 to end of array ($CMD[0] contains the length of the array.
			RunWait('"' & @ComSpec & '" /c ' & 'assoc ' & $CMD[$i], @SystemDir, @SW_HIDE)
		Next ;Next continues the loop
	EndIf ;EndIf ends this if statement
EndFunc ;Ends Function

Func OpenWorkgroup() ;Opens advanced computer options and opens the dialog to change workgroup and computer name settings
	RunWait('"' & @ComSpec & '" /c ' & 'control.exe %windir%\system32\sysdm.cpl', @SystemDir,@SW_HIDE) ;pulls up Advanced system properties
	ControlClick("System Properties","&Change...",115,"primary",1) ;Opens change workgroup dialog
EndFunc ;Ends Function