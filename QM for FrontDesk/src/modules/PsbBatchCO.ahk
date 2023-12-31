#Include "../../../Lib/Classes/utils.ahk"
#Include "../../../Lib/QM for FrontDesk/reports.ahk"

class PsbBatchCO {
    static name := "PsbBatchCO"
    static description := "旅安系统批量退房 - Excel表：CheckOut.xls"
    static popupTitle := "PSB CheckOut(Batch)"
    static scriptHost := "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\QM2 - Nightly"
    static path := IniRead(this.scriptHost . "\Lib\QM for FrontDesk\config.ini", "PsbBatchCO", "xlsPath")
    static blockingPopups := []

    static USE(desktopMode := 0) {
        if (desktopMode = true) {
            if (FileExist(A_Desktop . "\CheckOut.xls")) {
                path := A_Desktop . "\CheckOut.xls"
            } else if (FileExist(A_Desktop . "\CheckOut.xlsx")) {
                path := A_Desktop . "\CheckOut.xlsx"
            } else {
                MsgBox("对应 Excel表：CheckOut.xls并不存在！`n 请先创建或复制文件到桌面！", this.popupTitle)
                return
            }
        } else {
            path := this.path
        }
        quitOnRoom := IniRead(A_ScriptDir . "\src\lib\config.ini", "PsbBatchCO", "errorQuitAt")
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
        selector := MsgBox(textMsg, this.popupTitle, "YesNoCancel")
        if (selector = "No") {
            this.batchOut(path)
        } else if (selector = "Yes") {
            this.saveDep(path)
            this.batchOut(path)
        } else {
            utils.cleanReload(winGroup)
        }
    }

    static saveDep(path) {
        timePeriod := InputBox(
            textMsg := "
        (
        请选择FO03 - Departures数据时间范围：
    
        1 - 夜班： 00:00 ~ 07:00
        2 - 早班： 07:00 ~ 15:00
        3 - 中班： 15:00 ~ 23:59
        0 - 自定义时间段
        )", this.popupTitle)
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
                frTime := InputBox("请输入开始时间（格式：“hhmm”）", this.popupTitle).Value
                toTime := InputBox("请输入结束时间（格式：“hhmm”）", this.popupTitle).Value
            default:
                MsgBox("请输入对应时间的指令。")
        }

        BlockInput true
        WinMaximize "ahk_class SunAwtFrame"
        WinActivate "ahk_class SunAwtFrame"
        ; ; reportOpenDelimitedData("departure_all")
        ; Sleep 100
        ; MouseMove 473, 570
        ; Sleep 100
        ; Click
        ; MouseMove 474, 547
        ; Sleep 100
        ; Click
        ; MouseMove 483, 453
        ; Sleep 100
        ; Click
        ; MouseMove 613, 443
        ; Sleep 100
        ; Click
        ; Sleep 300
        ; MouseMove 495, 366
        ; Sleep 100
        ; Click "Down"
        ; MouseMove 378, 365
        ; Sleep 100
        ; Click "Up"
        ; Sleep 100
        ; Send "{Text}" . frTime
        ; Sleep 100
        ; MouseMove 579, 366
        ; Sleep 100
        ; Click "Down"
        ; MouseMove 449, 366
        ; Sleep 100
        ; Click "Up"
        ; Sleep 100
        ; Send "{Text}" . toTime
        ; Sleep 100
        ; roomNumsTxt := Format("{1} check-out.txt", today)
        ; ; reportSave(roomNumsTxt)
        ; ; }

        ; TODO: change the following part to loopfield and write xls
        ; A_Clipboard := FileRead(Format("{1}\{2}", A_MyDocuments, roomNumsTxt))
        BlockInput false
        Run path
        WinWait "ahk_class XLMAIN"
        WinMaximize "ahk_class XLMAIN"
        Sleep 3500
        WinWait "ahk_class QWidget"
        WinActivate "ahk_class QWidget"
        Sleep 500
        Send "{E}"
        Sleep 1000
        MouseMove 225, 962
        Sleep 100
        Click
        MouseMove 22, 207
        Sleep 100
        Click
        Sleep 100
        Send "{Delete}"
        Sleep 100
        Send "^v"
        Sleep 300
        MouseMove 930, 213
        Sleep 100
        Click
        Sleep 100
        Send "^x"
        Sleep 300
        MouseMove 149, 962
        Sleep 100
        Click
        MouseMove 70, 226
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
        MsgBox(continueText, this.popupTitle)
    }

    static batchOut(path) {
        autoOut := MsgBox("即将开始自动拍Out脚本`n请先打开PSB，进入“旅客退房”界面", PsbBatchCo.popupTitle, "OKCancel")
        if (autoOut = "Cancel") {
            utils.cleanReload(winGroup)
        }
        Xl := ComObject("Excel.Application")
        CheckOut := Xl.Workbooks.Open(path)
        depRooms := CheckOut.Worksheets("Sheet1")
        lastRow := depRooms.Cells(depRooms.Rows.Count, "A").End(-4162).Row
        depRoomNums := []
        TrayTip "读入房号中……"
        loop lastRow {
            depRoomNums.Push(Integer(depRooms.Cells(A_Index, 1).Text))
        }
        CheckOut.Close()
        Xl.Quit()
        TrayTip "读入完成"

        SetTimer(this.winDetectAndClose(this.blockingPopups))

        loop lastRow {
            A_Clipboard := depRoomNums[A_Index]
            MouseMove 279, 200
            Sleep 50
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
            MouseMove 279, 200
            Sleep 800
            Click "Down"
            MouseMove 160, 200

            Sleep 200
            Click "Up"
            Sleep 50
            Send "{BackSpace}"
            Sleep 200
            ; terminate on error pop-up
            if (WinExist("出错会导致溢出的弹窗名字")) {
                MsgBox("PSB系统出错，脚本已终止`n`n已拍Out到：" . depRoomNums[A_Index], PsbBatchCo.popupTitle)
                quitOnRoom := depRoomNums[A_Index]
                IniWrite(quitOnRoom, this.scriptHost . "\Lib\QM for FrontDesk\config.ini", "PsbBatchCO", "errorQuitAt")
                utils.cleanReload(winGroup)
            }
        }

        IniWrite("null", this.scriptHost . "\Lib\QM for FrontDesk\config.ini", "PsbBatchCO", "errorQuitAt")
        Sleep 1000
        MsgBox("PSB 批量拍Out 已完成！", this.popupTitle)

        SetTimer(this.winDetectAndClose(this.blockingPopups), 0)
    }

    static winDetectAndClose(targetWins) {
        loop targetWins.Length {
            if (WinExist(winGroup[A_Index])) {
                WinClose winGroup[A_Index]
            }
        }
    }
}