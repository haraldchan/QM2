; { import
; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\modules\CityLedgerCO.ahk"
; #Include "%A_ScriptDir%\src\modules\InhShare.ahk"
; #Include "%A_ScriptDir%\src\modules\Pbpf.ahk"
; #Include "%A_ScriptDir%\src\modules\GroupShare.ahk"
; #Include "%A_ScriptDir%\src\modules\DoNotMove.ahk"
; #Include "%A_ScriptDir%\src\modules\ReportMaster.ahk"
; #Include "%A_ScriptDir%\src\modules\GroupKeys.ahk"
; #Include "%A_ScriptDir%\src\modules\GroupProfilesModify.ahk"
; #Include "%A_ScriptDir%\src\modules\PsbBatchCO.ahk"
; #Include "%A_ScriptDir%\src\modules\Phrases.ahk"
#Include "./src/lib/utils.ahk"
#Include "./src/modules/CityLedgerCO.ahk"
#Include "./src/modules/InhShare.ahk"
#Include "./src/modules/Pbpf.ahk"
#Include "./src/modules/GroupShare.ahk"
#Include "./src/modules/DoNotMove.ahk"
#Include "./src/modules/ReportMaster.ahk"
#Include "./src/modules/GroupKeys.ahk"
#Include "./src/modules/GroupProfilesModify.ahk"
#Include "./src/modules/PsbBatchCO.ahk"
#Include "./src/modules/Phrases.ahk"
;}
; { setup
#SingleInstance Force
TraySetIcon A_ScriptDir . "\src\assets\QMTray.ico"
TrayTip "QM 2 运行中…按下 F9 开始使用脚本"
CoordMode "Mouse", "Screen"
; globals and states
version := "2.1.0"
today := FormatTime(A_Now, "yyyyMMdd")
config := A_ScriptDir . "\src\lib\config.ini"
popupTitle := "QM for FrontDesk " . version
cityLedgerOn := true
desktopMode := false
; script classes
scriptIndex := [
    [
        InhShare,
        Pbpf,
        GroupShare,
        DoNotMove,
        ReportMaster,
    ],
    [
        GroupKeys,
        GroupProfilesModify,
        PsbBatchCO,
    ],
]
; admin/Opera running detection
try {
    WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
} catch {
    MsgBox("
    (
    当前为非管理员状态或未运行Opera PMS，将无法正确运行。
    请在QM 2 图标上右键点击，选择“以管理员身份运行”；
    及打开Opera PMS。   
    )", popupTitle)
    ExitApp
}
; }

; { GUI template
QM := Gui("", popupTitle)
QM.AddText(, "
(
快捷键及对应功能：

F9:        显示脚本选择窗
F12:       强制停止脚本/重载
Ctrl+F12:  退出

常驻脚本(按下即启动)
Ctrl+O:    CityLedger挂账
)")
QM.AddCheckbox("Checked h25 y+10", "令 CityLedger 挂账保持常驻").OnEvent("Click", cityLedgerKeepAlive)

tab3 := QM.AddTab3("w350", ["一键运行", "Excel辅助", "常用语句"])
tab3.UseTab(1)
basic := []
loop scriptIndex[1].Length {
    radioStyle := (A_Index = 1) ? "Checked h25" : "h25 y+10"
    basic.Push(
        QM.AddRadio(radioStyle, scriptIndex[1][A_Index].description),
    )
}
reportMasterInfo := "
(
Report Master 常见问题：


因报表保存过程中出现弹窗可能会导致中断，建议：

1、 使用 IE 浏览器进行操作；

2、 将登陆页面最小化；

3、 重启浏览器以及 QM 2
)"
QM.AddText("y+25", reportMasterInfo)

tab3.UseTab(2)
xldp := []
loop scriptIndex[2].length {
    radioStyle := (A_Index = 1) ? "Checked h20 y+10" : "x20 h20 y+10"
    xldp.Push(
        [
            QM.AddRadio(radioStyle, scriptIndex[2][A_Index].description),
            QM.AddEdit("h25 w150 x20 y+10", scriptIndex[2][A_Index].path),
            QM.AddButton("h25 w70 x+20", "选择文件"),
            QM.AddButton("h25 w70 x+10", "打开表格"),
        ]
    )
}
loop xldp.Length {
    xldp[A_Index][1].OnEvent("Click", singleSelect.Bind(xldp[A_Index][1]))
    xldp[A_Index][2].OnEvent("LoseFocus", saveXlPath.Bind(scriptIndex[2][A_Index].name, xldp[A_Index][2]))
    xldp[A_Index][3].OnEvent("Click", getXlPath.Bind(scriptIndex[2][A_Index].name, xldp[A_Index][2]))
    xldp[A_Index][4].OnEvent("Click", openXlFile.Bind(xldp[A_Index][2].Text))
}
QM.AddCheckbox("h25 x20 y+10", "使用桌面文件模式").OnEvent("Click", toggleDesktopMode)
xldpInfo := "
(
功能说明：

启动脚本前，必须先将对应的数据从 Opera PMS 导出的报

表中复制到 Excel 表中，才能实现功能。


桌面文件模式：

选中“使用桌面文件模式”后，脚本将只会从本机桌面读取相

应文件名的 Excel 表。 请直接在桌面操作所需 Excel 表。
)"
QM.AddText("y+25", xldpInfo)

tab3.UseTab(3)
Phrases.USE(QM)

tab3.UseTab() ; end tab3

QM.AddButton("Default h40 w165", "启动脚本").OnEvent("Click", runSelectedScript)
QM.AddButton("h40 w165 x+18", "隐藏窗口").OnEvent("Click", (*) => QM.Hide())
; }

; { function scripts
cityLedgerKeepAlive(*) {
    global cityLedgerOn := !cityLedgerOn
}

singleSelect(ctrlObj, *) {
    loop xldp.Length {
        xldp[A_Index][1].Value := 0
    }
    ctrlObj.Value := 1
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

toggleDesktopMode(*) {
    global desktopMode := !desktopMode
    loop xldp.Length {
        xldp[A_Index][2].Enabled := !desktopMode
        xldp[A_Index][3].Enabled := !desktopMode
        xldp[A_Index][4].Enabled := !desktopMode
    }
}

runSelectedScript(*) {
    if (tab3.Value = 1) {
        loop basic.Length {
            if (basic[A_Index].Value = 1) {
                QM.Hide()
                scriptIndex[1][A_Index].USE()
            }
        }
    } else if (tab3.Value = 2) {
        loop xldp.Length {
            if (xldp[A_Index][1].Value = 1) {
                QM.Hide()
                scriptIndex[2][A_Index].USE(desktopMode)
            }
        }
    } else {
        return
    }
}
; }

; hotkey setup
F9:: {
    QM.Show()
 } 
F12:: cleanReload()
^F12:: quitApp() 

#HotIf cityLedgerOn
^o::CityLedgerCo.USE()
MButton::CityLedgerCo.USE()
#Hotif WinActive(popupTitle)
Esc:: {
    QM.Hide()
}
Enter:: runSelectedScript()