; #Include "%A_ScriptDir%\lib\utils.ahk"
; #Include "%A_ScriptDir%\config.ini"
#Include "../lib/utils.ahk"
#Include "../config.ini"
; date RegEx
dateDash := "^\d{1,4}-\d{1,2}-\d{1,2}"
dateSlash := "^\d{1,4}/\{1,2}/d{1,2}"

; TODO:Create GUI: create a file select ui
GroupKeysMain() {
    startMsg := "
    (
    即将进行批量团卡制作，启动前请先完成以下准备工作：

    1、请先保存需要制卡的团Rooming List；
    2、将Rooming List的房号录入“GroupKeys.xls”文件的第一列；
     - 如需单独修改某个房间的退房日期、时间，请分别填入GroupKeys.xls的第二、第三列
     - 日期格式：yyyy-mm-dd 或 yyyy/mm/dd（具体请先查看VingCard中格式）
     - 时间格式：HH:MM
    3、确保VingCard已经打开处于Check-in界面。
    )"

    start := MsgBox(startMsg, "GroupKeys", "OKCancel 4096")
    if (start = "Cancel") {
        cleanReload()
    }
    coDateInput := InputBox("请输入退房日期：", "GroupKeys").Value
    loop {
        if (RegExMatch(coDateInput, dateDash) > 0
            || RegExMatch(coDateInput, dateSlash)
            || coDateInput = "") {
            break
        } else {
            MsgBox("输入格式必须为yyyy-mm-dd或yyyy/mm/dd，如2023-01-01或2023/01/01。")
        }
    }
    coTimeInput := InputBox("请输入退房时间：", "GroupKeys", , "13:00").Value
    infoConfirm := MsgBox(Format("
    (
    当前团队制卡信息：
    退房日期：{1}
    退房时间：{2}
    )", coDateInput, coTimeInput), "GroupKeys", "OKCancel")
    if (infoConfirm = "Cancel") {
        cleanReload()
    }
    path := IniRead("%A_ScriptDir%\config.ini", "GroupKeys", "xlsPath")
    Xl := ComObject("Excel.Application")
    Xlbook := Xl.Workbooks.Open(path)
    groupRooms := Xlbook.Worksheets("Sheet1")
    lastRow := Xlbook.ActiveSheet.UsedRange.Rows.Count
    roomNum := Integer(groupRooms.Cells(1, 1))
    coDateRead := Integer(groupRooms.Cells(1, 2))
    coTimeRead := Integer(groupRooms.Cells(1, 3))
    A_Clipboard := roomNum
    row := 1

    BlockInput true
    loop lastRow {
        coDateLoop := coDateInput
        coTimeLoop := coTimeInput
        ; { paste room number (y-pos already modified!)
        MouseMove 421, 392
        Sleep 300
        Click "Down"
        MouseMove 255, 395
        Sleep 150
        Click "Up"
        MouseMove 387, 379
        Sleep 150
        Click 2
        Sleep 78
        MouseMove 395, 395
        Sleep 150
        Send "^v"
        Sleep 200
        ; }
        if (coDateRead != "") {
            coDateLoop := coDateRead
        }
        if (coTimeRead != "") {
            coTimeLoop := coTimeRead
        }
        MouseMove 417, 558
        Sleep 150
        Click "Down"
        MouseMove 233, 564
        Sleep 150
        Click "Up"
        Sleep 100
        Send coDateLoop
        Sleep 100
        MouseMove 524, 557
        Sleep 100
        Click 2
        Sleep 100
        Send coTimeLoop
        Sleep 100
        MouseMove 505, 737
        Sleep 150
        Click "Down"
        MouseMove 505, 738
        Sleep 93
        Click "Up"
        Sleep 100
        Send "2"
        Sleep 100
        Send "!e"
        Sleep 100

        check := MsgBox(Format("
        (
        已做房卡：{1}
         - 是(Y)制作下一个
         - 否(N)退出制卡
        )", roomNum), "GroupKeys","OKCancel 4096")
        if (check = "Cancel") {
            cleanReload()
        }
        row += 1
        roomNum := Integer(groupRooms.Cells(row, 1))
        A_Clipboard := roomNum
        coDateRead := Integer(groupRooms.Cells(row, 2))
        coTimeRead := Integer(groupRooms.Cells(row, 3))
    }
    BlockInput false
    Xlbook.Close
    Xl.Quit
    MsgBox("已完成团队制卡，请与Opera/蓝豆系统核对是否正确！","GroupKeys")
}


; hotkeys
; !F1:: GroupKeysMain()
; F12:: cleanReload()	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
