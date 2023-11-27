; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"

class GroupKeys {
    static name := "GroupKeys"
    static description := "团队房卡制作     - Excel表：GroupKeys.xls"
    static popupTitle := "Group Keys"
    static dateDash := "^\d{1,4}-\d{1,2}-\d{1,2}"
    static dateSlash := "\d{4}/\d{2}/\d{2}"
    static path := IniRead(A_ScriptDir . "\src\lib\config.ini", "GroupKeys", "xlsPath")

    static USE(desktopMode := 0) {
        if (desktopMode = true) {
            if (FileExist(A_Desktop . "\GroupKeys.xls")) {
                path := A_Desktop . "\GroupKeys.xls"
            } else if (FileExist(A_Desktop . "\GroupKeys.xlsx")) {
                path := A_Desktop . "\GroupKeys.xlsx"
            } else {
                MsgBox("对应 Excel表：GroupKeys.xls并不存在！`n 请先创建或复制文件到桌面！", this.popupTitle)
                return
            }
        } else {
            path := this.path
        }
        start := MsgBox("
        (
        即将进行批量团卡制作，启动前请先完成以下准备工作：

        1、请先保存需要制卡的团Rooming List；

        2、将Rooming List的房号录入“GroupKeys.xls”文件的第一列；

        - 如需单独修改某个房间的退房日期、时间，请分别填入GroupKeys.xls的第二、第三列
        - 日期格式：yyyy-mm-dd 或 yyyy/mm/dd（具体请先查看VingCard中格式）
        - 时间格式：HH:MM
        
        3、确保VingCard已经打开处于Check-in界面。
        )", this.popupTitle, "OKCancel 4096")
        if (start = "Cancel") {
            cleanReload()
        }
        loop {
            coDateInput := InputBox("请输入退房日期：`n(格式为yyyy-mm-dd或yyyy/mm/dd)", this.popupTitle).Value
            if (RegExMatch(coDateInput, this.dateDash) > 0
                || RegExMatch(coDateInput, this.dateSlash)
                || coDateInput = "") {
                    break
            } else {
                MsgBox("输入格式必须为yyyy-mm-dd或yyyy/mm/dd，如2023-01-01或2023/01/01。")
                continue
            }
        }
        coTimeInput := InputBox("请输入退房时间：`n(格式为HH:MM)", this.popupTitle, , "13:00").Value

        infoConfirm := MsgBox(Format("
            (
            当前团队制卡信息：
            退房日期：{1}
            退房时间：{2}
            )", coDateInput, coTimeInput), "GroupKey", "OKCancel")
        if (infoConfirm = "Cancel") {
            cleanReload()
        }
        
        Xl := ComObject("Excel.Application")
        GroupKeysXl := Xl.Workbooks.Open(path)
        groupRooms := GroupKeysXl.Worksheets("Sheet1")
        lastRow := groupRooms.Cells(groupRooms.Rows.Count, "A").End(-4162).Row

        roomNums := []
        coDateRead := []
        coTimeRead := []
        loop lastRow {
            roomNums.Push(groupRooms.Cells(A_Index, 1).Text)
            groupRooms.Cells(A_Index, 2).Text = "" 
                ? coDateRead.Push("blank") 
                : coDateRead.Push(groupRooms.Cells(A_Index, 2).Text)
            groupRooms.Cells(A_Index, 3).Text = "" 
                ? coTimeRead.Push("blank")
                : coTimeRead.Push(groupRooms.Cells(A_Index, 3).Text)
        }
        GroupKeysXl.Close()
        Xl.Quit()

        loop lastRow {
            BlockInput true
            A_Clipboard := roomNums[A_Index]
            coDateLoop := (coDateRead[A_Index] = "blank") ? coDateInput : coDateRead[A_Index]
            coTimeLoop := (coTimeRead[A_Index] = "blank") ? coTimeInput : coTimeRead[A_Index]
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
                )", roomNums[A_Index]), this.popupTitle, "OKCancel 4096")
            if (checkConf = "Cancel") {
                cleanReload()
            }
        }
        Sleep 1000
        MsgBox("已完成团队制卡，请与Opera/蓝豆系统核对是否正确！", this.popupTitle)
    }
}