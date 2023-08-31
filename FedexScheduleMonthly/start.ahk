/* TODO: create a GUI
    1. how-to info
    2. file select -> path
    3. schedule date range -> range
*/
; #Include "%A_ScriptDir\FedexScheduleMonthly.ahk%"
#Include ../FedexScheduleMonthly.ahk
; configs
config := Format("{1}\config.ini", A_ScriptDir)
path := IniRead(config, "FSM", "schedulePath")
schdRange := IniRead(config, "FSM", "scheduleRange")

; { GUI
FSM := Gui("+Resize", "Fedex Schedule Monthly")
FSM.AddText("x20 y20", "请选择Schedule 文件") 

; schedule path (by input or file select)
schdPath := FSM.AddEdit("h25 w150 x20 y+10", path) 
schdPath.OnEvent("LoseFocus", savePath)
FSM.AddButton("h25 w70 x+20", "选择文件").OnEvent("Click", getSchd) 

; determine date range as part of foldername
FSM.AddText("x20 y+20", "请选择Schedule 覆盖日期区间(将成为文件夹后缀)",)
rangeCtrl := FSM.AddEdit("h25 w150 x20 y+10", schdRange)
rangeCtrl.OnEvent("LoseFocus", saveRange)

; click to run main script
; FSM.AddButton("h20 w50 x20 y+20", "开始生成").OnEvent("Click", FsmMain())
beginBtn := FSM.AddButton("h25 w70 y+20", "开始生成").OnEvent("Click",genSchd)
FSM.AddButton("h25 w70 x+20", "退出脚本").OnEvent("Click", quit)

FSM.Show()
; }

; { callbacks 
savePath(*) {
    if (schdPath.Text = "") {
        MsgBox("请选择 Schedule 文件")
        return
    }
    IniWrite(schdPath.Text, config, "FSM", "schedulePath")
}

saveRange(*) {
    if (rangeCtrl.Text = "") {
        MsgBox("请选择输入日期范围")
        return
    }
    IniWrite(rangeCtrl.Text, config, "FSM", "scheduleRange")
}

getSchd(*) {
    FSM.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Schedule 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    schdPath.Value := selectFile
    IniWrite(selectFile, config, "FSM", "schedulePath")
}

genSchd(*) {
    if (schdPath.Text = "") {
        MsgBox("请选择 Schedule 文件")
        return
    }
    if (schdPath.Text = "") {
        MsgBox("请选择 Schedule 文件")
        return
    }
    FsmMain()
}

quit(*) {
    ExitApp
}
; }

; util functions
cleanReload(){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    ; Key/Mouse state set default
    BlockInput false
	SetCapsLockState false
    CoordMode "Mouse", "Client"
    ; Excel Quit
    ComObject("Excel.Application").Quit()
	Reload
}


;hotkeys
F12:: cleanReload()