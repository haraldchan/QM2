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
FSM := Gui("+Resize W450 H200", "Fedex Schedule Monthly")
FSM.AddText(, "请选择Schedule 文件") 
; schedule path (by input or file select)
schdPath := FSM.AddEdit("h20 w100 x20 y20", path) 
schdPath.OnEvent("LoseFocus", savePath)
FSM.AddButton("h20 w50 x+20", "选择文件").OnEvent("Click", getSchedule) 
; determine date range as part of foldername
FSM.AddText("y+20", "请选择Schedule 覆盖日期区间(将成为文件夹后缀)",)
rangeCtrl := FSM.AddEdit("h20 w100 x20 y+20", schdRange).OnEvent("LoseFocus", saveRange)
FSM.Add() ; divider
; click to run main script
FSM.AddButton("h20 w50 x20 y+20", "开始生成").OnEvent("Click", FsmMain())
FSM.AddButton("h20 w50 x+20", "退出脚本").OnEvent("Click", (*) => ExitApp)

FSM.Show()
; }

; { callbacks 
savePath(*) {
    IniWrite(schdPath.Value, config, "FSM", "schedulePath")
}

saveRange(*) {
    IniWrite(rangeCtrl.Value, config, "FSM", "scheduleRange")
}

getSchedule(*) {
    FSM.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Schedule 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    schdPath.Value := selectFile
    IniWrite(selectFile, config, "FSM", "schedulePath")
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