; reminder: Y-pos needs to minus 20
; #Include "%A_ScriptDir%\lib\utils.ahk"
#Include "../lib/utils.ahk"

GroupProfilesModifyMain() {
    path := IniRead("%A_ScriptDir%\config.ini", "GroupProfilesModify", "xlsPath")
    wwly := IniRead("%A_ScriptDir%\config.ini", "GroupProfilesModify", "wwlyPath")

    gpmStart := MsgBox("
    (
    将开始团队Profile modify

    注意事项：
    1、请确保同一个房间内，PSB与Opera客人数量一致（关注一个人、三个人住）；
    2、请在InHouse界面下启动；
    3、请确保旅业信息同步系统已经启动。
    )", "GroupProfilesModify", "OKCancel 4096")
    if (gpmStart = "Cancel") {
        cleanReload()
    }
    Xl := ComObject("Excel.Application")
    Xlbook := Xl.Workbooks.Open(path)
    groupRooms := Xlbook.Worksheets("Sheet1")
    lastRow := Xlbook.ActiveSheet.UsedRange.Rows.Count
    roomNum := Integer(groupRooms.Cells(1, 1))
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
        WinWait "" ;TODO: use title instead of class!

        MouseMove 690, 243
        Sleep 1300
        Click
        Sleep 100
        Click "Down"
        MouseMove 575, 240
        Sleep 450
        Click "Up"
        Sleep 200
        Send "^v"
        Sleep 1000
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

        ; TODO: find out hex color of the error popup
        ; errorBrown := ""
        ; if (PixelGetColor(251, 196) = errorBrown) {
        ;     MsgBox(Format("Modify出错，脚本已终止`n`n已Modify到：{1}", roomNum))
        ;     quitOnRoom := roomNum
        ;     IniWrite(quitOnRoom, "%A_ScriptDir%\config.ini", "PsbBatchCO", "errorQuitAt")
        ;     cleanReload()
        ; }
        row++
        roomNum := Integer(groupRooms.Cells(row, 1))
    }
    Xlbook.Close
    Xl.Quit
    BlockInput false
    MsgBox("已Modify 完成，请再次检查是否正确。", "GroupProfilesModify")
}


; hotkeys
; F4:: SharePbPfMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script