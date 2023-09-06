; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"
; date RegEx
; dateDash := "^\d{1,4}-\d{1,2}-\d{1,2}"
; dateSlash := "^\d{1,4}/\{1,2}/d{1,2}"
; scoped vals
; GroupKeys := {
;     popupTitle: "Group Keys"
; }

; GroupKeysMain() {
;     startMsg := "
;     (
;     即将进行批量团卡制作，启动前请先完成以下准备工作：

;     1、请先保存需要制卡的团Rooming List；
;     2、将Rooming List的房号录入“GroupKeys.xls”文件的第一列；
;      - 如需单独修改某个房间的退房日期、时间，请分别填入GroupKeys.xls的第二、第三列
;      - 日期格式：yyyy-mm-dd 或 yyyy/mm/dd（具体请先查看VingCard中格式）
;      - 时间格式：HH:MM
;     3、确保VingCard已经打开处于Check-in界面。
;     )"
;     start := MsgBox(startMsg, GroupKeys.popupTitle, "OKCancel 4096")
;     if (start = "Cancel") {
;         cleanReload()
;     }
;     loop {
;         coDateInput := InputBox("请输入退房日期：", GroupKeys.popupTitle).Value
;         if (RegExMatch(coDateInput, dateDash) > 0
;             || RegExMatch(coDateInput, dateSlash)
;             || coDateInput = "") {
;                 break
;         } else {
;             MsgBox("输入格式必须为yyyy-mm-dd或yyyy/mm/dd，如2023-01-01或2023/01/01。")
;             continue
;         }
;     }
;     coTimeInput := InputBox("请输入退房时间：", GroupKeys.popupTitle, , "13:00").Value

;     infoConfirm := MsgBox(Format("
;     (
;     当前团队制卡信息：
;     退房日期：{1}
;     退房时间：{2}
;     )", coDateInput, coTimeInput), "GroupKey", "OKCancel")
;     if (infoConfirm = "Cancel") {
;         cleanReload()
;     }
;     path := IniRead(config, "GroupKeys", "xlsPath")
;     Xl := ComObject("Excel.Application")
;     GroupKeysXl := Xl.Workbooks.Open(path)
;     groupRooms := GroupKeysXl.Worksheets("Sheet1")
;     lastRow := groupRooms.Cells(groupRooms.Rows.Count, "A").End(-4162).Row
;     row := 1

;     loop lastRow {
;         BlockInput true
;         roomNum := groupRooms.Cells(row, 1).Text
;         coDateRead := groupRooms.Cells(row, 2).Text
;         coTimeRead := groupRooms.Cells(row, 3).Text
;         A_Clipboard := roomNum
;         coDateLoop := (coDateRead = "") ? coDateInput : coDateRead
;         coTimeLoop := (coTimeRead = "") ? coTimeInput : coTimeRead
;         ; { paste room number (y-pos already modified!)
;         MouseMove 387, 409
;         Sleep 300
;         Click "Down"
;         MouseMove 252, 409
;         Sleep 150
;         Click "Up"
;         Sleep 150
;         Send "^v"
;         Sleep 200
;         MouseMove 410, 582
;         Sleep 150
;         Click "Down"
;         MouseMove 249, 582
;         Sleep 150
;         Click "Up"
;         Sleep 100
;         Send "{Text}" . coDateLoop
;         Sleep 100
;         MouseMove 528, 578
;         Sleep 150
;         Click 2
;         Sleep 200
;         Send "{Text}" . coTimeLoop
;         Sleep 100
;         MouseMove 499, 742
;         Sleep 100
;         Click 2
;         Sleep 100
;         Send "{Text}2"
;         Sleep 100
;         Send "!e"
;         Sleep 100
;         BlockInput false
;         checkConf := MsgBox(Format("
;         (
;         已做房卡：{1}
;          - 是(Y)制作下一个
;          - 否(N)退出制卡
;         )", roomNum), GroupKeys.popupTitle, "OKCancel 4096")
;     if (checkConf = "Cancel") {
;         cleanReload()
;     }
;     row++
;     }
;     GroupKeysXl.Close
;     Xl.Quit
;     MsgBox("已完成团队制卡，请与Opera/蓝豆系统核对是否正确！", GroupKeys.popupTitle)
; }

class GroupKeys {
    static popupTitle := "Group Keys"
    static dateDash := "^\d{1,4}-\d{1,2}-\d{1,2}"
    static dateSlash := "^\d{1,4}/\{1,2}/d{1,2}"

    static Main() {
        start := MsgBox("
        (
        即将进行批量团卡制作，启动前请先完成以下准备工作：

        1、请先保存需要制卡的团Rooming List；
        2、将Rooming List的房号录入“GroupKeys.xls”文件的第一列；
        - 如需单独修改某个房间的退房日期、时间，请分别填入GroupKeys.xls的第二、第三列
        - 日期格式：yyyy-mm-dd 或 yyyy/mm/dd（具体请先查看VingCard中格式）
        - 时间格式：HH:MM
        3、确保VingCard已经打开处于Check-in界面。
        )", GroupKeys.popupTitle, "OKCancel 4096")
        if (start = "Cancel") {
            cleanReload()
        }
        loop {
            coDateInput := InputBox("请输入退房日期：", this.popupTitle).Value
            if (RegExMatch(coDateInput, this.dateDash) > 0
                || RegExMatch(coDateInput, this.dateSlash)
                || coDateInput = "") {
                    break
            } else {
                MsgBox("输入格式必须为yyyy-mm-dd或yyyy/mm/dd，如2023-01-01或2023/01/01。")
                continue
            }
        }
        coTimeInput := InputBox("请输入退房时间：", this.popupTitle, , "13:00").Value

        infoConfirm := MsgBox(Format("
            (
            当前团队制卡信息：
            退房日期：{1}
            退房时间：{2}
            )", coDateInput, coTimeInput), "GroupKey", "OKCancel")
        if (infoConfirm = "Cancel") {
            cleanReload()
        }
        path := IniRead(config, "GroupKeys", "xlsPath")
        Xl := ComObject("Excel.Application")
        GroupKeysXl := Xl.Workbooks.Open(path)
        groupRooms := GroupKeysXl.Worksheets("Sheet1")
        lastRow := groupRooms.Cells(groupRooms.Rows.Count, "A").End(-4162).Row
        row := 1

        loop lastRow {
            BlockInput true
            roomNum := groupRooms.Cells(row, 1).Text
            coDateRead := groupRooms.Cells(row, 2).Text
            coTimeRead := groupRooms.Cells(row, 3).Text
            A_Clipboard := roomNum
            coDateLoop := (coDateRead = "") ? coDateInput : coDateRead
            coTimeLoop := (coTimeRead = "") ? coTimeInput : coTimeRead
            ; { paste room number (y-pos already modified!)
            MouseMove 387, 409
            Sleep 300
            Click "Down"
            MouseMove 252, 409
            Sleep 150
            Click "Up"
            Sleep 150
            Send "^v"
            Sleep 200
            MouseMove 410, 582
            Sleep 150
            Click "Down"
            MouseMove 249, 582
            Sleep 150
            Click "Up"
            Sleep 100
            Send "{Text}" . coDateLoop
            Sleep 100
            MouseMove 528, 578
            Sleep 150
            Click 2
            Sleep 200
            Send "{Text}" . coTimeLoop
            Sleep 100
            MouseMove 499, 742
            Sleep 100
            Click 2
            Sleep 100
            Send "{Text}2"
            Sleep 100
            Send "!e"
            Sleep 100
            BlockInput false
            checkConf := MsgBox(Format("
                (
                已做房卡：{1}
                - 是(Y)制作下一个
                - 否(N)退出制卡
                )", roomNum), GroupKeys.popupTitle, "OKCancel 4096")
        if (checkConf = "Cancel") {
            cleanReload()
        }
        row++
        }
        GroupKeysXl.Close
        Xl.Quit
        MsgBox("已完成团队制卡，请与Opera/蓝豆系统核对是否正确！", GroupKeys.popupTitle)
    }
}