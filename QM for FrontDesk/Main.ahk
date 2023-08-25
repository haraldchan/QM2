; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\SharePbPf.ahk"
; #Include "%A_ScriptDir%\src\CityLedgerCO.ahk"
; #Include "%A_ScriptDir%\src\ReportMaster.ahk"
; #Include "%A_ScriptDir%\src\GroupShareDNM.ahk"
; #Include "%A_ScriptDir%\src\PsbBatchCO.ahk"
; #Include "%A_ScriptDir%\src\GroupKeys.ahk"
; #Include "%A_ScriptDir%\src\GroupProfilesModify.ahk"
#Include "../src/lib/utils.ahk"
#Include "../src/SharePbPf.ahk"
#Include "../src/CityLedgerCO.ahk"
#Include "../src/ReportMaster.ahk"
#Include "../src/GroupShareDNM.ahk"
#Include "../src/PsbBatchCO.ahk"
#Include "../src/GroupKeys.ahk"
#Include "../src/GroupProfilesModify.ahk"

TrayTip "QM 2 运行中。 按下Alt+F1 可查看快捷键指南"
tips := "
(
快捷键:		脚本功能:

F9: 		In-house Share & PayBy/PayFor
F4: 		批量DoNotMove & 旅行团(TGDA)房批量Share+DNM
Ctrl+O:		CityLedger CheckOut
Ctrl+F11:		常用报表保存
Ctrl+F1:        团队房卡批量制作
Ctrl+F4:        团队Profile 批量录入
Ctrl+F9:        旅安系统批量旅客退房（拍Out）

F12:		刷新脚本状态
Ctrl+F12:		退出

备注： 部分功能可能会收到输入法干扰，建议使用前切换到英文输入法。
)"

; universal miscs
!F1:: MsgBox(tips, "QM for FrontDesk 2.0.0")
F12:: cleanReload()	; use 'Reload' for script reset
^F12:: ExitApp	; use 'ExitApp' to kill script

; Script hotkeys
#Hotif WinExist("ahk_class SunAwtFrame") 
F9:: SharePbPfMain()
^o:: CityLedgerCOMain()
^F11:: ReportMasterMain()
F4:: GroupShareDnmMain()
^F9:: PsbBatchCoMain()
^F1:: GroupKeysMain()
^F4:: GroupProfilesModifyMain()
