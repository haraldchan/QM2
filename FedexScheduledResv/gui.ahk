; { imports
; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\FedexScheduledResv.ahk"
#Include "../utils.ahk"
#Include "../FedexScheduledResv.ahk"
; }
; { setup
#SingleInstance Force
; check "Run as admin", ExitApp if false
try {
    WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
} catch {
    MsgBox("
    (
    当前为非管理员状态，将无法正确运行。
    请在FSR 图标上右键点击，选择“以管理员身份运行”。    
    )", "Fedex Scheduled Reservations")
    ExitApp
}
config := A_ScriptDir . "\config.ini"
path := IniRead(config, "FSR", "schedulePath")
toNextDayTime := IniRead(config, "FSR", "toNextDayTime")
; }
; { GUI
FSR := Gui("+Resize", "Fedex Scheduled Reservations")
FSR.AddText(, "
(
快捷键及对应功能：

F12:       强制停止脚本
Ctrl+F12:  退出

)")
FSR.AddText("x20 y+20", "(请选择Schedule 文件")
; schedule path (by input or file select)
schdPath := FSR.AddEdit("h25 w150 x20 y+10", path)
schdPath.OnEvent("LoseFocus", savePath)
FSR.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getSchd)

;TODO: add dropdown list to determine toNextDayTime

; click to run main script / quit FSR
FSR.AddButton("h25 w70 y+20", "开始录入").OnEvent("Click", runFSR)
FSR.AddButton("h25 w70 x+20", "退出脚本").OnEvent("Click", quit)

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