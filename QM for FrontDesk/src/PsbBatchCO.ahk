; #Include "%A_ScriptDir%\lib\utils.ahk"
; #Include "%A_ScriptDir%\lib\reports.ahk"
#Include "../lib/utils.ahk"
#Include "../lib/reports.ahk"

; TODO:Create GUI: create a file select ui
today := FormatTime(A_Now, "yyyyMMdd")
config := Format("{1}\src\config.ini", A_ScriptDir)
popupTitle := "PSB CheckOut(Batch)"
path := IniRead(config, "PsbBatchCO", "xlsPath")

PsbBatchCoMain() {
    quitOnRoom := IniRead(config, "PsbBatchCO", "errorQuitAt")
    if (quitOnRoom != "null") {
        MsgBox(Format("上次拍Out为出错停止，已拍至：{1}`n请先更新CheckOut.xls", quitOnRoom))
    }

    textMsg := "
    (
    即将开始批量拍Out 脚本

    是(Y) = 保存FO03 - Departures数据txt 文件后开始
    否(N) = 直接开始拍Out
    取消 = 退出脚本
    )"

    selector := MsgBox(textMsg, popupTitle, "YesNoCancel")
    if (selector = "No") {
        batchOut()
    } else if (selector = "Yes") {
        saveDep()
        batchOut()
    } else {
        cleanReload()
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
    timePeriod := InputBox(textMsg, popupTitle)
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
            frTime := InputBox("请输入开始时间（格式：“hhmm”）", popupTitle).Value
            toTime := InputBox("请输入结束时间（格式：“hhmm”）", popupTitle).Value
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
    Send Format("{Text}{1}", frTime)
    Sleep 100
    MouseMove 579, 346
    Sleep 100
    Click "Down"
    MouseMove 449, 346
    Sleep 100
    Click "Up"
    Sleep 100
    Send Format("{Text}{1}", toTime)
    Sleep 100
    roomNumsTxt := Format("{1} check-out.txt", today)
    reportSave(roomNumsTxt)
    ; }
    A_Clipboard := FileRead(Format("{1}\{2}", A_MyDocuments, roomNumsTxt))
    BlockInput false
    Run path
    WinWait "ahk_class XLMAIN"
    WinMaximize "ahk_class XLMAIN"
    Sleep 3500
    ; { pasting data (y-pos already modified!)
    WinWait "ahk_class QWidget"
    WinActivate "ahk_class QWidget"
    Sleep 500
    Send "{E}"
    Sleep 1000
    MouseMove 225, 942
    Sleep 100
    Click
    MouseMove 22, 187
    Sleep 100
    Click
    Sleep 100
    Send "{Delete}"
    Sleep 100
    Send "^v"
    Sleep 300
    MouseMove 930, 193
    Sleep 100
    Click
    Sleep 100
    Send "^x"
    Sleep 300
    MouseMove 149, 942
    Sleep 100
    Click
    MouseMove 70, 206
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
    MsgBox(continueText, popupTitle)
}

batchOut() {
    autoOut := MsgBox("即将开始自动拍Out脚本`n请先打开PSB，进入“旅客退房”界面", popupTitle, "OKCancel")
    if (autoOut := "Cancel") {
        cleanReload()
    }
    Xl := ComObject("Excel.Application")
    CheckOut := Xl.Workbooks.Open(path)
    depRooms := CheckOut.Worksheets("Sheet1")
    lastRow := depRooms.Cells(depRooms.Rows.Count,"A").End(-4162).Row
    MsgBox(lastRow)
    row := 1
    errorBrown := "0x804000" ; TODO: confirm this hex color code!
    ; BlockInput true
    CoordMode "Mouse", "Screen"
    loop lastRow {
        roomNum := Integer(depRooms.Cells(row, 1).Value)
        A_Clipboard := roomNum
        ; { check out in psb (y-pos already modified!)
        MouseMove 279, 204
        Sleep 500
        Click
        Sleep 350
        Send "^v"
        Sleep 100
        Send "^f"
        Sleep 800
        Send "{Enter}"
        Sleep 800
        Send "{Enter}"
        Sleep 800
        MouseMove 279, 204
        Sleep 800
        Click "Down"
        MouseMove 160, 204
        Sleep 200
        Click "Up"
        Sleep 50
        Send "{BackSpace}"
        Sleep 200
        ; }
        ; terminate on error pop-up
        if (PixelGetColor(251, 196) = errorBrown) {
            MsgBox(Format("PSB系统出错，脚本已终止`n`n已拍Out到：{1}", roomNum))
            quitOnRoom := roomNum
            IniWrite(quitOnRoom, config, "PsbBatchCO", "errorQuitAt")
            cleanReload()
        }
        row++
    }
    ; BlockInput false
    CheckOut.Close
    Xl.Quit
    IniWrite("null", config, "PsbBatchCO", "errorQuitAt")
    MsgBox("PSB 批量拍Out 已完成！", "PSB CheckOut(Batch)")
}


; hotkeys
; ^F9:: PsbBatchCoMain()
; ^F9:: saveDep()
; F12:: cleanReload()   ; use 'Reload' for script reset
; ^F12:: ExitApp    ; use 'ExitApp' to kill script

