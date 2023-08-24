; #Include "%A_ScriptDir%\Lib\reports.ahk"
#Include "../Lib/reports.ahk"
#Include "../config.ini"

/* TODO:Create GUI
    1. create a file select ui
    2. create ini file for file path(for xls depencies) store
*/
today := FormatTime(A_Now, "yyyyMMdd")
PsbBatchCoMain() {
    textMsg := "
    (
    即将开始批量拍Out 脚本

    是(Y) = 保存FO03 - Departures数据txt 文件后开始
	否(N) = 直接开始拍Out
	取消 = 退出脚本
    )"

    selector := MsgBox(textMsg, "PSB CheckOut(Batch)", "YesNoCancel")
    if (selector = "No") {
        pakOut()
    } else if (selector = "Yes") {
        saveDep()
        pakOut()
    } else {
        Reload
    }
}

saveDep() {
    textMsg := "
    (
    请选择FO03 - Departures数据时间范围：

    1 - 夜班： 00:00 ~ 07:00
    2 - 早班： 07:00 ~ 15:00
    3 - 中班： 15:00 ~ 23:59
    0 - 自定义时间段
    )"
    timePeriod := InputBox(textMsg, "PSB CheckOut(Batch)")
    if (timePeriod.Result = "Cancel") {
        Reload
    }
    switch timePeriod.Value {
        Case "1":
            frTime := "0000"
            toTime := "0700"
        Case "2":
            frTime := "0600"
            toTime := "1500"
        Case "3":
            frTime := "1400"
            toTime := "2359"
        Case "0":
            frTime := InputBox("请输入开始时间（格式：“hhmm”）","PSB CheckOut(Batch)").Value
            toTime := InputBox("请输入结束时间（格式：“hhmm”）","PSB CheckOut(Batch)").Value
        default:
            MsgBox("请输入对应时间的指令。")
    }

    BlockInput true
    WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"
    ; { saving departured room numbers (y-pos already modified!)
    reportOpenDelimitedData("departure_all")
    Sleep 100
    MouseMove 473, 550
    Sleep 100
    Click
    MouseMove 474, 527
    Sleep 100
    Click
    MouseMove 483, 433
    Sleep 100
    Click
    MouseMove 613, 423
    Sleep 100
    Click
    Sleep 300
    MouseMove 495, 346
    Sleep 100
    Click "Down"
    MouseMove 378, 345
    Sleep 100
    Click "Up"
    Sleep 100
    Send frTime
    Sleep 100
    MouseMove 579, 346
    Sleep 100
    Click "Down"
    MouseMove 449, 346
    Sleep 100
    Click "Up"
    Sleep 100
    Send toTime
    Sleep 100
    roomNumsTxt := Format("{1} check-out.txt", today)
    reportSave(roomNumsTxt)
    ; }
    MsgBox("保存房号中，请等待(5)秒","PSB CheckOut(Batch)","T5 4096")
    A_Clipboard := FileRead("%A_MyDocuments%\%roomNumsTxt%")
    BlockInput false

    ; TODO: open excel file and paste
    path := IniRead("%A_ScriptDir%\config.ini", "PsbBatchCO", "xlsPath")
    Run path
    WinWait "" ;TODO: findout wps's ahk_class 
    WinMaximize "" 
    Sleep 3500
    ; { pasting data (y-pos already modified!)
    MouseMove 735, 408
	Sleep 500
    Send "{E}"
	Sleep 1000
	MouseMove 231, 921
	Sleep 100
	Click
	MouseMove 70, 195
	Sleep 100
	Click
	Sleep 100
    Send "^v"
    Sleep 300
	MouseMove 946, 177
	Sleep 100
	Click
	Sleep 100
	Send "^x"
    Sleep 300
	MouseMove 176, 918
	Sleep 100
	Click
	MouseMove 91, 174
	Sleep 100
	Click
	Sleep 100
    Send "^v"
    Sleep 300
    ; }
    continueText := "
    (
    请打开蓝豆查看“续住工单”，剔除excel中续住的房间后，点击确定继续拍out。
    完成上述操作前请不要按掉此弹窗。
    )"
    MsgBox(continueText, "PSB CheckOut(Batch)")
}

pakOut() {

}

; hotkeys
; ^F9:: SharePbPfMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
