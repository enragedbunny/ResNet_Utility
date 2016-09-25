#include <RunCMD.au3>
local $iShowDialog = 1
local $iCloseDialog = 0
local $sData = "ipconfig/all"

_RunCMD($sData,$iShowDialog,$iCloseDialog) 