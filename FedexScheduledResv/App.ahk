; { imports
; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\FedexScheduledResv.ahk"
#Include "../utils.ahk"
#Include "../FedexScheduledResv.ahk"
; }
; { setup
#SingleInstance Force
TraySetIcon A_ScriptDir . "\assets\FSRTray.ico"
try {
    WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
} catch {
    MsgBox("
    (
    当前为非管理员状态或未运行Opera PMS，将无法正确运行。
    请在QM 2 图标上右键点击，选择“以管理员身份运行”；
    及打开Opera PMS。 
    )", "Fedex Scheduled Reservations")
    ExitApp
}
config := A_ScriptDir . "\config.ini"
path := IniRead(config, "FSR", "schedulePath")
resvOnDayTime := IniRead(config, "FSR", "resvOnDayTime")
; }

; { GUI
FSR := Gui("+Resize", "Fedex Scheduled Reservations")
FSR.AddText(, "
(
快捷键及对应功能：

F12:       强制停止脚本

)")
tab3 := FSR.AddTab3("w500 h350", ["选项", "启动界面参考"])
tab3.UseTab(1)
FSR.AddText("x20 y+20", "1. 请选择Schedule 文件")
; schedule path (by input or file select)
schdPath := FSR.AddEdit("h25 w150 x20 y+10", path)
schdPath.OnEvent("LoseFocus", savePath)
FSR.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getSchd)
FSR.AddButton("h25 w70 x+20", "打开文件").OnEvent("Click", openSchd)

;TODO: add dropdown list to determine resvOnDayTime
FSR.AddText("x20 y+30 ", "2. 请指定提前一天留房时间点：`n（ETA在此时间点后的房间不会提前一天留房）")
toNextDropdown := FSR.AddDropDownList("w150 x20 y+10 Choose2", ["09:00", "10:00", "11:00", "12:00", "13:00"])
toNextDropdown.OnEvent("Change", saveTime)

; click to run main script / quit FSR
FSR.AddButton("h25 w70 y+20", "开始录入").OnEvent("Click", runFSR)
FSR.AddButton("h25 w70 x+20", "退出脚本").OnEvent("Click", quit)

tab3.UseTab(2)
FSR.AddPicture("w500 h-1", A_ScriptDir . "\assets\RsvnList.JPG")

FSR.Show()
; }

; { callback funcs
savePath(*) {
    IniWrite(schdPath.Text, config, "FSR", "schedulePath")
}

getSchd(*) {
    FSR.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Schedule 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    schdPath.Value := selectFile
    IniWrite(selectFile, config, "FSR", "schedulePath")
}

openSchd(*) {
    Run schdPath.Text
}

saveTime(*) {
    IniWrite(toNextDropdown.Text, config, "FSR", "resvOnDayTime")
}

runFSR(*) {
    if (schdPath.Text = "") {
        MsgBox("请选择 Schedule 文件")
        return
    }
    FSR.Hide()
    FsrMain()
}

quit(*) {
    ExitApp
}
; }
; hotkey setups
F12:: cleanReload()
^F12:: quitApp()