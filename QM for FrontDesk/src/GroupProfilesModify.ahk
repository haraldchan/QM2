; reminder: Y-pos needs to minus 20
; #Include "%A_ScriptDir%\lib\utils.ahk"
#Include "../lib/utils.ahk"
popupTitle := "GroupProfilesModify"

GroupProfilesModifyMain() {
    path := IniRead("%A_ScriptDir%\config.ini", popupTitle, "xlsPath")
    wwly := IniRead("%A_ScriptDir%\config.ini", popupTitle, "wwlyPath")
    errorRed := "0x800000"
    gpmStart := MsgBox("
    (
    将开始团队Profile modify

    注意事项：
    1、请确保同一个房间内，PSB与Opera客人数量一致（关注一个人、三个人住）；
    2、请在InHouse界面下启动；
    3、请确保旅业信息同步系统已经启动。
    )", popupTitle, "OKCancel 4096")
    if (gpmStart = "Cancel") {
        cleanReload()
    }
    Xl := ComObject("Excel.Application")
    GroupRoomNum := Xl.Workbooks.Open(path)
    groupRooms := GroupRoomNum.Worksheets("Sheet1")
    lastRow := GroupRoomNum.ActiveSheet.Cells(GroupRoomNum.ActiveSheet.Rows.Count,"A").End(-4162).Row
    roomNum := Integer(groupRooms.Cells(1, 1).Text)
    row := 1

    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    A_Clipboard := roomNum
    BlockInput true
    loop lastRow {
        MouseMove 598, 553
        Sleep 500
        Send "!p"
        Sleep 300
        Send "!n"
        Sleep 200
        MouseMove 666, 525
        Sleep 1000
        Click
        Sleep 3000
        Run wwly
        Send "{Pause}"
        WinWait "ahk_class TMAIN"
        WinActivate "ahk_class TMAIN"

        ; MouseMove 690, 243
        ; Sleep 1300
        ; Click
        ; Sleep 100
        ; Click "Down"
        ; MouseMove 575, 240
        ; Sleep 450
        ; Click "Up"
        ; Sleep 200
        ; Send "^v"
        ; Sleep 1000
        ; Send "{F11}"
        ; Sleep 300

        ; window coord
        MouseMove 400, 23
        Click "Down"
        Sleep 200
        MouseMove 324, 23
        Click "Up"
        Sleep 200
        Send A_Clipboard
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

        if (PixelGetColor(698, 306) = errorRed) {
            MsgBox(Format("Modify出错，脚本已终止`n`n已Modify到：{1}", roomNum))
            quitOnRoom := roomNum
            IniWrite(quitOnRoom, "%A_ScriptDir%\config.ini", "PsbBatchCO", "errorQuitAt")
            cleanReload()
        }
        row++
        roomNum := Integer(groupRooms.Cells(row, 1).Value)
        Sleep 500
    }
    GroupRoomNum.Close
    Xl.Quit
    BlockInput false
    MsgBox("已Modify 完成，请再次检查是否正确。", popupTitle)
}


; hotkeys
; F4:: SharePbPfMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
