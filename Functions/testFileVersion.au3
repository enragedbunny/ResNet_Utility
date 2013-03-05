local $sVer = FileGetVersion(@ProgramFilesDir & "\Internet Explorer\iexplore.exe")
local $sVer1 = FileGetVersion(@ProgramFilesDir & "\Mozilla Firefox\firefox.exe")
msgbox(0,"title",$sVer & " " & $sVer1)