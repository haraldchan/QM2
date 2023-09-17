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
; check "Run as admin", ExitApp if false
try {
    WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
} catch {
    MsgBox("
    (
    当前为非管理员状态或未运行Opera PMS，将无法正确运行。
    请在QM 2 图标上右键点击，选择“以管理员身份运行”；
    及打开Opera PMS。   
    )", "QM for FrontDesk 2.1.0")
    ExitApp
}
TraySetIcon A_ScriptDir . "\src\assets\QMTray.ico"
TrayTip "QM 2 运行中…按下 F9 开始使用脚本"
CoordMode "Mouse", "Screen"
today := FormatTime(A_Now, "yyyyMMdd")
config := A_ScriptDir . "\src\config.ini"
cityLedgerOn := false
desktopMode := false
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
QM := Gui("+Resize", "QM for FrontDesk 2.1.0")
QM.AddText(, "
(
快捷键及对应功能：

F9:        显示脚本选择窗
F12:       强制停止脚本
Ctrl+F12:  退出

常驻脚本(按下即启动)
Ctrl+O:    CityLedger挂账

)")
QM.AddCheckbox("h25 y+10", "令 CityLedger 挂账保持常驻").OnEvent("Click", cityLedgerKeepAlive)

tab3 := QM.AddTab3("w350 h300", ["基础功能", "Excel辅助", "ReportMaster"])
tab3.UseTab(1)
basic := [
    QM.AddRadio("Checked h20 y+10", "空白InHouse Share"),
    QM.AddRadio("h25 y+10", "粘贴PayBy PayFor"),
    QM.AddRadio("h25 y+10", "旅行团Share + DoNotMove"),
    QM.AddRadio("h25 y+10", "批量DoNotMove Only"),
]
QM.AddButton("Default h25 w70 x30 y420", "启动脚本").OnEvent("Click", runSelectedScript.Bind(1))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)

tab3.UseTab(2)
xldp := [
    [
        gk := QM.AddRadio("Checked h20 y+10", "团队房卡制作     - Excel表：GroupKeys.xls"),
        gk.OnEvent("Click", singleSelect.Bind(gk)),
        gkXl := QM.AddEdit("h25 w150 x20 y+10", GroupKeys.path),
        gkBtn1 := QM.AddButton("h25 w70 x+20", "选择文件"),
        gkBtn2 := QM.AddButton("h25 w70 x+10", "打开表格"),
    ],
    [
        gpm := QM.AddRadio("h20 x20 y+10", "团队Profile录入  - Excel表：GroupRoomNum.xls"),
        gpm.OnEvent("Click", singleSelect.Bind(gpm)),
        gpmXl := QM.AddEdit("h25 w150 x20 y+10", GroupProfilesModify.path),
        gpmBtn1 := QM.AddButton("h25 w70 x+20", "选择文件"),
        gpmBtn2 := QM.AddButton("h25 w70 x+10", "打开表格"),
    ],
    [
        co := QM.AddRadio("h20 x20 y+10", "旅业系统批量退房 - Excel表：CheckOut.xls"),
        co.OnEvent("Click", singleSelect.Bind(co)),
        coXl := QM.AddEdit("h25 w150 x20 y+10", PsbBatchCO.path),
        coBtn1 := QM.AddButton("h25 w70 x+20", "选择文件"),
        coBtn2 := QM.AddButton("h25 w70 x+10", "打开表格"),
    ],
]
; { xldp events
gkXl.OnEvent("LoseFocus", saveXlPath.Bind("GroupKeys", gkXl))
gkBtn1.OnEvent("Click", getXlPath.Bind("GroupKeys", gkXl))
gkBtn2.OnEvent("Click", openXlFile.Bind(gkXl.Text))
gpmXl.OnEvent("LoseFocus", saveXlPath.Bind("GroupProfilesModify", gpmXl))
gpmBtn1.OnEvent("Click", getXlPath.Bind("GroupProfilesModify", gpmXl))
gpmBtn2.OnEvent("Click", openXlFile.Bind(gpmXl.Text))
coXl.OnEvent("LoseFocus", saveXlPath.Bind("PsbBatchCO", coXl))
coBtn1.OnEvent("Click", getXlPath.Bind("PsbBatchCO", coXl))
coBtn2.OnEvent("Click", openXlFile.Bind(coXl.Text))
; }
desktopModeCheck := QM.AddCheckbox("vDesktopMode h25 x20 y+15", "使用桌面文件模式").OnEvent("Click", toggleDesktopMode)
QM.AddButton("Default h25 w70 x30 y420", "启动脚本").OnEvent("Click", runSelectedScript.Bind(2))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)

tab3.UseTab(3)
QM.AddText("h20", "`n点击“启动脚本”打开报表选择器.")
QM.AddButton("Default h25 w70 x30 y420", "启动脚本").OnEvent("Click", runSelectedScript.Bind(3))
QM.AddButton("h25 w70 x+20", "隐藏窗口").OnEvent("Click", hideWin)
; }

; { callback funcs
cityLedgerKeepAlive(*) {
    global cityLedgerOn := !cityLedgerOn
}

toggleDesktopMode(*) {
    global desktopMode := !desktopMode
    gkXl.Enabled := !desktopMode
    gkBtn1.Enabled := !desktopMode
    gkBtn2.Enabled := !desktopMode
    gpmXl.Enabled := !desktopMode
    gpmBtn1.Enabled := !desktopMode
    gpmBtn2.Enabled := !desktopMode
    coXl.Enabled := !desktopMode
    coBtn1.Enabled := !desktopMode
    coBtn2.Enabled := !desktopMode
}

runSelectedScript(currentTab, *) {
    if (currentTab = 1) {
        loop basic.Length {
            if (basic[A_Index].Value = 1) {
                hideWin()
                if (A_Index = 1) {
                    SharePbPf.share()
                } else if (A_Index = 2) {
                    SharePbPf.pbpf()
                } else if (A_Index = 3) {
                    GroupShareDNM.dnmShare()
                } else if (A_Index = 4) {
                    GroupShareDNM.dnm()
                }
            }
        }
    } else if (currentTab = 2) {
        loop xldp.Length {
            if (xldp[A_Index][1].Value = 1) {
                hideWin()
                scriptIndex[2][A_Index].Main(desktopMode)
            }
        }
    } else {
        hideWin()
        scriptIndex[3].Main()
    }
}

saveXlPath(script, ctrlObj, *) {
    IniWrite(ctrlObj.Text, config, script, "xlsPath")
}

getXlPath(script, ctrlObj, *) {
    QM.Opt("+OwnDialogs")
    selectedFile := FileSelect(3, , "请选择 Excel 文件")
    if (selectedFile = "") {
        MsgBox("请选择文件")
        return
    }
    ctrlObj.Value := selectedFile
    IniWrite(selectedFile, config, script, "xlsPath")
}

openXlFile(file, *) {
    Run file
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
F9:: {
    QM.Show()
 } ; show QM2 window
F12:: cleanReload()	; use 'Reload' for script reset
^F12:: quitApp() ; use ExitApp to kill app
#HotIf cityLedgerOn
^o::CityLedgerCo.Main()
#Hotif WinActive("ahk_class AutoHotkeyGUI")
Esc:: hideWin()
Enter:: runSelectedScript(tab3.Value)