; #Include "%A_ScriptDir%\src\SharePbPf.ahk"
; #Include "%A_ScriptDir%\src\CityLedgerCO.ahk"
; #Include "%A_ScriptDir%\src\ReportMaster.ahk"
; #Include "%A_ScriptDir%\src\GroupShareDNM.ahk"

#Include "../src/SharePbPf.ahk"
#Include "../src/CityLedgerCO.ahk"
#Include "../src/ReportMaster.ahk"
#Include "../src/GroupShareDNM.ahk"
TrayTip "QM 2 运行中。 按下Alt+F1 可查看快捷键指南"

tips := "
(
快捷键:		脚本功能:

F9: 		In-house Share & PayBy/PayFor
F4: 		批量DoNotMove & 旅行团(TGDA)房批量Share+DNM
Ctrl+O:		CityLedger CheckOut
Ctrl+F11:		常用报表保存

F12:		刷新脚本状态
Ctrl+F12:		退出

备注： 部分功能可能会收到输入法干扰，建议使用前切换到英文输入法。
)"

cleanReload(){
	Reload
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	SetCapsLockState false
}

; universal miscs
F12:: cleanReload()	; use 'Reload' for script reset
^F12:: ExitApp	; use 'ExitApp' to kill script
!F1:: MsgBox(tips, "QM for FrontDesk 2.0.0")

; Script hotkeys
#Hotif WinExist("ahk_class SunAwtFrame") 
F9:: SharePbPfMain()
F4:: GroupShareDnmMain()
^o:: CityLedgerCOMain()
^F11:: ReportMasterMain()
