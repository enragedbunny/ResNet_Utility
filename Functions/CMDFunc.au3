;This file contains many functions for doing various things from command line
;I don't think they depend on any other files
;The FixFileAssociations function is probably the best example of how we should do these (if you want to add other functions to this file).


Func AltGetMAC()
	RunWait('"' & @ComSpec & '" /k ' & 'ipconfig /all', @SystemDir,@SW_SHOW) ; RunWait starts the indicated program and waits until it loads before executing other code.
EndFunc

Func ResetNetwork() 
	local $i
	local $Data = 'netsh winsock reset|' & _
				 'netsh winsock reset catalog|' & _
				 'netsh interface ip reset c:\int-resetlog.txt|' & _
				 'netsh interface ip delete arpcache|' & _
				 'ipconfig /flushdns|' & _
				 'ipconfig /registerdns'
	local $CMD = StringSplit($Data,"|")
	
	If IsArray($CMD) Then ;checks if $CMD is an array then runs the loop.
		For $i = 1 to $CMD[0] ;looping from 1 to the end of the array because CMD[0] is the length from string split
				RunWait('"' & @ComSpec & '" /c ' & $CMD[$i], @SystemDir,@SW_HIDE) ;@ComSpec is cmd prompt, and /c is for closing it. @SW_HIDE is a autoit command that hides the window, starts it hidden. @SystemDir is the working directory (not the path to the file).
		Next
	EndIf  
EndFunc

Func AddRemovePrograms()
	RunWait('"' & @ComSpec & '" /c ' & 'appwiz.cpl', @SystemDir,@SW_HIDE) ;types appwiz.cpl in cmd prompt: runs add remove programs
EndFunc

Func RepairFirewall()
   local $CMD1 = 'rundll32.exe setupapi.dll,InstallHinfSection Ndi-Steelhead 132 %windir%\inf\netrass.inf' ;changing the initialization

   RunWait('"' & @ComSpec & '" /c ' & $CMD1, @SystemDir,@SW_HIDE)
EndFunc

Func ResetFirewall()
   local $CMD1 = 'netsh firewall reset'

   RunWait('"' & @ComSpec & '" /c ' & $CMD1, @SystemDir,@SW_HIDE)
EndFunc

Func RepairWinUpdate()
   Local $i
   Local $Data1 = 'net stop bits|' & _ ;& _ continues the next line... eg: net stop bits|net stop wuauserv|...
			     'net stop wuauserv|' & _
			     'Del "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat"|' & _
				 'Ren %systemroot%\SoftwareDistrobution\DataStore *.bak|' & _
				 'Ren %systemroot%\SoftwareDistrobution\Download *.bak|' & _
				 'Ren %systemroot%\system32\catroot2 *.bak|' & _
				 'sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|' & _
				 'sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|' & _
				 'cd /d %WINDIR%\system32'
   Local $Data2 = 'atl.dll|' & _
				 'urlmon.dll|' & _
				 'mshtml.dll|' & _
				 'shdocvw.dll|' & _
				 'browseui.dll|' & _
				 'jscript.dll|' & _
				 'vbscript.dll|' & _
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
   Local $Data3 = 'netsh reset winsock|' & _
				  'net start bits|' & _
				  'net start wuauserv'
   Local $CMD = StringSplit($Data1, "|") ;Converts Data1 into an array $CMD
   Local $CMDReg = StringSplit($Data2, "|")
   Local $CMD2 = StringSplit($Data3, "|")
   If IsArray($CMD) Then ;checks if $CMD is an array then runs the loop.
	  For $i = 1 to $CMD[0] ;looping from 1 to the end of the array because CMD[0] is the length from string split
		 RunWait('"' & @ComSpec & '" /c ' & $CMD[$i], @SystemDir, @SW_HIDE) ;This loopp is running all of the functions in Data1.
	  Next ;Next continues the loop
   EndIf ;EndIf ends this if statement
;so once this loop is closed it will run this next data set vvvvv
   If IsArray($CMDReg) Then
	  For $i = 1 to $CMDReg[0]
		 RunWait('"' & @ComSpec & '" /c ' & 'regsvr32.exe /s ' & $CMDReg[$i], @SystemDir, @SW_HIDE)
	  Next
   EndIf

   If IsArray($CMD2) Then
	  For $i = 1 to $CMD2[0]
		 RunWait('"' & @ComSpec & '" /c ' & $CMD2[$i], @SystemDir, @SW_HIDE)
	  Next
   EndIf
EndFunc ;Ends Function

Func FixFileAssociations() ;needs to be expanded a ton
   Local $i
   Local $Data = '.exe=exefile|' & _
			     '.bat=batfile|' & _
				 '.reg=regfile|' & _
				 '.com=comfile|' & _
				 '.xml=xmlfile|' & _
				 '.lnk=lnkfile|' & _
				 '.ico=icofile|'
   Local $CMD = StringSplit($Data, "|") ;String split saves length as CMD[0] in the array.
   If IsArray($CMD) Then
	  For $i = 1 to $CMD[0]
		 RunWait('"' & @ComSpec & '" /c ' & 'assoc ' & $CMD[$i], @SystemDir, @SW_HIDE)
	  Next
   EndIf
EndFunc

Func OpenWorkgroup()
	RunWait('"' & @ComSpec & '" /c ' & 'control.exe %windir%\system32\sysdm.cpl', @SystemDir,@SW_HIDE) ;pulls up control panel>
	ControlClick("System Properties","&Change...",115,"primary",1) ;System>changeSettings
EndFunc