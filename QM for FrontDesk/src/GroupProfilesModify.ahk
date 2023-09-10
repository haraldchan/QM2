; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"

class GroupProfilesModify {
    static popupTitle := "GroupProfilesModify"
    static path := IniRead(A_ScriptDir . "\src\config.ini", "GroupProfilesModify", "xlsPath")
    static wwly := IniRead(A_ScriptDir . "\src\config.ini", "GroupProfilesModify", "wwlyPath")

    static Main() {
        errorRed := "0x800000"
        gpmStart := MsgBox("
        (
        将开始团队Profile modify
    
        注意事项：
        1、请确保同一个房间内，PSB与Opera客人数量一致（关注一个人、三个人住）
        2、请在InHouse界面下启动
        )", GroupProfilesModify.popupTitle, "OKCancel 4096")
        if (gpmStart = "Cancel") {
            cleanReload()
        }
        Xl := ComObject("Excel.Application")
        GroupRoomNum := Xl.Workbooks.Open(this.path)
        groupRooms := GroupRoomNum.Worksheets("Sheet1")
        lastRow := groupRooms.Cells(groupRooms.Rows.Count,"A").End(-4162).Row
        roomNum := Integer(groupRooms.Cells(1, 1).Text)
        row := 1
        Run this.wwly
        Sleep 5000
        WinMaximize "ahk_class SunAwtFrame"
        WinActivate "ahk_class SunAwtFrame"
        A_Clipboard := roomNum
        BlockInput true
        loop lastRow {
            MouseMove 598, 573
            Sleep 500
            Send "!p"
            Sleep 300
            Send "!n"
            Sleep 200
            MouseMove 666, 545
            Sleep 1000
            Click
            Sleep 3000

            Send "{Pause}"
            WinWait "ahk_class TMAIN"
            WinActivate "ahk_class TMAIN"
    
            CoordMode "Mouse", "Client"
            MouseMove 400, 23
            Click "Down"
            Sleep 200
            MouseMove 324, 23
            Click "Up"
            Sleep 200
            Send "^v"
            Sleep 200
            Send "{F11}"
            Sleep 300
            Send "{Enter}"
            Sleep 7500
            MouseMove 626, 299
            Sleep 500
            Send "!o"
            Sleep 500
            Send "!o"
            Sleep 2000
            Send "{Down}"
            Sleep 200
            CoordMode "Mouse", "Screen"
    
            if (PixelGetColor(698, 306) = errorRed) {
                MsgBox("Modify出错，脚本已终止`n`n已Modify到：" . roomNum, this.popupTitle)
                quitOnRoom := roomNum
                IniWrite(quitOnRoom, "%A_ScriptDir%\config.ini", "PsbBatchCO", "errorQuitAt")
                cleanReload()
            }
            row++
            try {
            roomNum := Integer(groupRooms.Cells(row, 1).Value)
            A_Clipboard := roomNum
            Sleep 500
            } catch {
                break
            }
        }
        GroupRoomNum.Close()
        Xl.Quit()
        BlockInput false
        MsgBox("已Modify 完成，请再次检查是否正确。", this.popupTitle)
    }
}