local $oCMDDict = ObjCreate("Scripting.Dictionary")
$oCMDDict.add("ipconfig","ipconfig /all|" & _   ;Shows Interface information
	          			 "1|" & _				;Shows dialog
	          			 "0")					;Leaves dialog
$oCMDDict.add("resetnetwork","netsh winsock reset|" & _                             ;Resets Winsock settings
				 			 "netsh winsock reset catalog|" & _                     ;Resets Winsock catalog
				 			 "netsh interface ip reset c:\int-resetlog.txt|" & _    ;Reset IP Interface settings and save a log
				 			 "netsh interface ip delete arpcache|" & _              ;Delete ARP cache on all IP interfaces
				 			 "ipconfig /flushdns|" & _                              ;Flush DNS
				 			 "ipconfig /registerdns|" & _                           ;Register DNS)
				 			 "0|" & _												;Hides dialog
				 			 "1")													;Closes dialog when finished
$oCMDDict.add("programsfeatures","appwiz.cpl|" & _	 ;Opens Programd and Features
								 "0|" & _			 ;Hides dialog
								 "1")			     ;Closes dialog when finished
$oCMDDict.add("repairfirewall","rundll32.exe setupapi.dll,InstallHinfSection Ndi-Steelhead 132 %windir%\inf\netrass.inf|" & _     ;Repairs Windows Firewall
							   "1|" & _																							  ;Shows dialog
							   "1")																								  ;Closes dialog when finished
$oCMDDict.add("resetfirewall","netsh firewall reset|" & _	;Resets Windows Firewall
			   				  "0|" & _						;Hides dialog
			   				  "1")							;Closes dialog when finished
$oCMDDict.add("repairwindowsupdate","net stop bits|" & _ 																	;Stop BITS service
			     					"net stop wuauserv|" & _   																;Stops wuauserv service
			     					'Del "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr*.dat"|' & _  ;Delete qmgr*.dat files
				 					"Ren %systemroot%\SoftwareDistrobution\DataStore *.bak|" & _ 							;Rename DataStore files/Folders to *.bak
				 					"Ren %systemroot%\SoftwareDistrobution\Download *.bak|" & _ 							;Rename Download Folder to Download.bak
				 					"Ren %systemroot%\system32\catroot2 *.bak|" & _ 										;Rename any files/folders that begin with catroot and change to *.bak
				 					"sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|" & _    ;Configure BITS settings (a windows service)
				 					"sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)|" & _ ;Configure WUAUServ settings (a windows service)
				 					"cd /d %WINDIR%\system32|" & _ 															;Changes working directory in cmd prompt to run the following DLL Register commands")
									"regsvr32.exe /s atl.dll|" & _
									"regsvr32.exe /s urlmon.dll|" & _
									"regsvr32.exe /s mshtml.dll|" & _
									"regsvr32.exe /s shdocvw.dll|" & _
									"regsvr32.exe /s browseui.dll|" & _
									"regsvr32.exe /s jscript.dll|" & _
									"regsvr32.exe /s vbscript.dll|" & _
									"regsvr32.exe /s scrrun.dll|" & _
									"regsvr32.exe /s msxml.dll|" & _
									"regsvr32.exe /s msxml3.dll||" & _
									"regsvr32.exe /s msxml6.dll|" & _
									"regsvr32.exe /s actxprxy.dll|" & _
									"regsvr32.exe /s softpub.dll|" & _
									"regsvr32.exe /s wintrust.dll|" & _
									"regsvr32.exe /s dssenh.dll|" & _
									"regsvr32.exe /s rsaenh.dll|" & _
									"regsvr32.exe /s gpkcsp.dll|" & _
									"regsvr32.exe /s sccbase.dll|" & _
									"regsvr32.exe /s slbcsp.dll|" & _
									"regsvr32.exe /s cryptdlg.dll|" & _
									"regsvr32.exe /s oleaut32.dll|" & _
									"regsvr32.exe /s ole32.dll|" & _
									"regsvr32.exe /s shell32.dll|" & _
									"regsvr32.exe /s initpki.dll|" & _
									"regsvr32.exe /s wuapi.dll|" & _
									"regsvr32.exe /s wuaueng.dll|" & _
									"regsvr32.exe /s wuaueng1.dll|" & _
									"regsvr32.exe /s wucltui.dll|" & _
									"regsvr32.exe /s wups.dll|" & _
									"regsvr32.exe /s wups2.dll|" & _
									"regsvr32.exe /s wuweb.dll|" & _
									"regsvr32.exe /s qmgr.dll|" & _
									"regsvr32.exe /s qmgrprxy.dll|" & _
									"regsvr32.exe /s wucltux.dll|" & _
									"regsvr32.exe /s muweb.dll|" & _
									"regsvr32.exe /s wuwebv.dll|" & _
									"netsh reset winsock|" & _                  ;Reset Winsock settings
									"net start bits|" & _                       ;Start BITS windows service
									"net start wuauserv|" & _ ;Start wuauserv windows service
									"0|" & _							     	;Hides dialog
									"1")										;Closes dialog when finished
$oCMDDict.add("fixextensions","assoc .exe=exefile|" & _   ;creates an association between .exe extensions and the program that opens exe Files
			     			  "assoc .bat=batfile|" & _   ;creates an association between .bat extensions and the program that opens bat Files
				 			  "assoc .reg=regfile|" & _   ;creates an association between .reg extensiosn and the program that opens reg Files
				 			  "assoc .com=comfile|" & _   ;creates an association between .com extensions and the program that opens com Files
				 			  "assoc .xml=xmlfile|" & _   ;creates an association between .xml extensions and the program that opens xml Files
				 			  "assoc .lnk=lnkfile|" & _   ;creates an association between .lnk (shortcuts) extensions and the program that opens lnk Files
				 			  "assoc .ico=icofile|" & _   ;creates an association between .ico (icons) extensions and the program that opens ico Files)
							  "0|" & _					  ;Hides dialog
							  "1")						  ;Closes dialog when finished
$oCMDDict.add("openworkgroup","control.exe %windir%\system32\sysdm.cpl|" & _	;Opens advanced computer properties
							  "0|" & _											;Hides dialog
							  "1")												;Closes dialog when finished