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
        GroupKeys,
        GroupProfilesModify,
        PsbBatchCO,
    ],
    ReportMaster
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

tab3 := QM.AddTab3("w300 h280", ["基础功能", "Excel辅助", "ReportMaster"])
tab3.UseTab(1)
basic := [
    QM.AddRadio("Checked h20 y+10", "空白InHouse Share"),
    QM.AddRadio("h20 y+10", "粘贴PayBy PayFor"),
    QM.AddRadio("h20 y+10", "旅行团Share + DoNotMove"),
    QM.AddRadio("h20 y+10", "批量DoNotMove Only"),
]
QM.AddButton("Default h25 w70 x30 y310", "启动脚本").OnEvent("Click", runSelectedScript.Bind(tab3.Value))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)

tab3.UseTab(2)
xldp := [
    [
    	gk := QM.AddRadio("Checked h20 y+10", "团队房卡制作     - Excel表：GroupKeys.xls"),
        gk.OnEvent("Click", singleSelect.Bind(gk)),
    	gkXl := QM.AddEdit("h25 w150 x20 y+10", GroupKeys.path),
    	gkXl.OnEvent("LoseFocus", saveXlPath.Bind("GroupKeys", gkXl)),
    	QM.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile.Bind("GroupKeys", gkXl)),
    ],
    [
    	gpm := QM.AddRadio("h20 x20 y+10", "团队Profile录入  - Excel表：GroupRoomNum.xls"),
        gpm.OnEvent("Click", singleSelect.Bind(gpm)),
    	gpmXl := QM.AddEdit("h25 w150 x20 y+10", GroupProfilesModify.path),
    	gpmXl.OnEvent("LoseFocus", saveXlPath.Bind("GroupProfilesModify", gpmXl)),
    	QM.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile.Bind("GroupProfilesModify", gpmXl)),
    ],
    [
    	co := QM.AddRadio("h20 x20 y+10", "旅业系统批量退房 - Excel表：CheckOut.xls"),
        co.OnEvent("Click", singleSelect.Bind(co)),
    	coXl := QM.AddEdit("h25 w150 x20 y+10", PsbBatchCO.path),
    	coXl.OnEvent("LoseFocus", saveXlPath.Bind("PsbBatchCO", coXl)),
    	QM.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getXlFile.Bind("PsbBatchCO", coXl)),
    ],
]
QM.AddButton("Default h25 w70 x30 y310", "启动脚本").OnEvent("Click", runSelectedScript.bind(tab3.Value))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)

tab3.UseTab(3)
QM.AddText("h20", "`n点击“启动脚本”打开报表选择器.")
QM.AddButton("Default h25 w70 x30 y310", "启动脚本").OnEvent("Click", runSelectedScript.bind(tab3.Value))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)
; }

; { callback funcs
runSelectedScript(currentTab, *) {
    if (currentTab = 1) {
        loop basic.Length {
            if (basic[A_Index].Value = 1) {
                QM.Hide()
                if (A_Index = 1) {
                    SharePbPf.share()
                } else if (A_Index = 2) {
                    SharePbPf.pbpf()
                } else if (A_Index = 3) {
                    GroupShareDNM.dnmShare()
                } else if (A_Index = 4) {
                    GroupShareDNM.dnm()
                }
                return
            }
        }
    } else if (currentTab = 2) {
        loop xldp.Length {
            if (xldp[A_Index][1].Value = 1) {
                QM.Hide()
                scriptIndex[2][A_Index].Main()
                return
            }
        }
    } else {
        QM.Hide()
        scriptIndex[3].Main()
    }
}

saveXlPath(script, ctrlObj, *) {
    IniWrite(ctrlObj.Text, config, script, "xlsPath")
}

getXlFile(script, ctrlObj, *) {
    QM.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Excel 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    ctrlObj.Value := selectFile
    IniWrite(selectFile, config, script, "xlsPath")
}

singleSelect(ctrlObj, *) {
    loop xldp.Length {
        xldp[A_Index][1].Value := 0
    }
    ctrlObj.Value := 1
}

hideWin(*) {
    QM.Hide()
}
; }

; hotkey setup
F9:: QM.Show() ; show QM2 window
F12:: cleanReload()	; use 'Reload' for script reset
^F12:: quitApp() ; use ExitApp to kill app

#Hotif WinExist("ahk_class AutoHotkeyGUI") 
Esc:: hideWin()
Enter:: runSelectedScript(tab3.Value,)