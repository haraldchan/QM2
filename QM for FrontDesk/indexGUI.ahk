; { import
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
;}
; { setup
#SingleInstance Force
TrayTip "QM 2 运行中…按下 F9 开始使用脚本"
CoordMode "Mouse", "Screen"
today := FormatTime(A_Now, "yyyyMMdd")
config := A_ScriptDir . "\src\config.ini"
scriptIndex := [
    [
        SharePbPf.share,
        SharePbPf.pbpf,
        GroupShareDNM.dnmShare,
        GroupShareDNM.dnm
    ],
    [
        GroupKeys.Main,
        GroupProfilesModify.Main,
        PsbBatchCO.Main
    ],
    ReportMaster.Main
]
; }

; { GUI
QM := Gui("+Resize", "QM for FrontDesk 2")
QM.AddText(, "
(
快捷键及对应功能：

F9:     显示脚本选择窗
F12:    强制停止脚本

)")

tab3 := QM.AddTab3("w300 h200", ["基础功能", "Excel辅助", "ReportMaster"])
tab3.UseTab(1)
basic := [
    QM.AddRadio("Checked h15 y+10", "空白InHouse Share")
    QM.AddRadio("h15 y+10", "粘贴PayBy PayFor")
    QM.AddRadio("h15 y+10", "旅行团Share + DoNotMove")
    QM.AddRadio("h15 y+10", "批量DoNotMove Only")
]
QM.AddButton("Default h25 w70 x30 y195", "启动脚本").OnEvent("Click", getSelectedScript(1))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", cleanReload)

tab3.UseTab(2)
xldp := [
    QM.AddRadio("Checked h15 y+10", "团队房卡制作     - Excel表：GroupKeys.xls")
    gkXl := QM.AddEdit("h25 w150 x20 y+10", GroupKeys.path)
    gkXl.OnEvent("LoseFocus", saveXlPath("GroupKeys", gkXl))
    gkXl.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile("GroupKeys", gkXl))
    QM.AddRadio("h15 y+10", "团队Profile录入  - Excel表：GroupRoomNum.xls")
    gpmXl := QM.AddEdit("h25 w150 x20 y+10", GroupProfilesModify.path)
    gpmXl.OnEvent("LoseFocus", saveXlPath("GroupProfilesModify", gpmXl))
    gpmXl.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile("GroupProfilesModify", gpmXl))
    QM.AddRadio("h15 y+10", "旅业系统批量退房 - Excel表：CheckOut.xls")
    coXl := QM.AddEdit("h25 w150 x20 y+10", PsbBatchCO.path)
    coXl.OnEvent("LoseFocus", saveXlPath("GroupProfilesModify", coXl))
    coXl.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile("GroupProfilesModify", coXl))
]
QM.AddButton("Default h25 w70 x30 y195", "启动脚本").OnEvent("Click", getSelectedScript(2))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", cleanReload)
; }

; { callback funcs
getSelectedScript(tab, *) {
    if (tab = 1) {
        loop basic.Length {
            if (basic[A_Index].Value = "On") {
                QM.Hide()
                scriptIndex[1][A_Index]()
                return
            }
        }
    } else if (tab = 2) {
        loop xldp.Length {
            if (xldp[A_Index].Value = "On") {
                QM.Hide()
                scriptIndex[2][A_Index]()
                return
            }
        }
    } else {
        QM.Hide()
        scriptIndex[3]()
    }
}

saveXlPath(script, control, *) {
    IniWrite(control.Text, config, script, "xlsPath")
}

getXlFile(script, control, *) {
    QM.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Excel 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    control.Value := selectFile
    IniWrite(selectFile, config, script, "xlsPath")
}
; }

; hotkey setup
F9:: QM.Show() ; show QM2 window
F12:: cleanReload()	; use 'Reload' for script reset
^F12:: quitApp() ; use ExitApp to kill app