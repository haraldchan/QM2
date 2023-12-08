; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "../lib/utils.ahk"
#Include  "../../../Lib/Classes/utils.ahk"

class GroupShare {
    static description := "旅行团房Share + DoNotMove"
    static popupTitle := "Group Share & DoNotMove"

    static USE() {
        WinMaximize "ahk_class SunAwtFrame"
        WinActivate "ahk_class SunAwtFrame"
        Sleep 500
        confirmMsg := MsgBox("
        (
        将开始批量团Share及DoNotMove操作。
    
        运行前请先确认：
            1.Opera窗口已最大化。
            2.界面在RoomAssign。
            3.以Name筛选团房（如使用BlockCode将会出错）
        )", this.popupTitle, "OKCancel 4096")
        If (confirmMsg = "Cancel") {
            Utils.cleanReload(winGroup)
        } else {
            roomQty := InputBox("`n请输入需要Share + DNM的房间数量", this.popupTitle, "h150")
            if (roomQty.Result = "Cancel") {
                Utils.cleanReload(winGroup)
            }
        }
        MouseMove 340, 311
        Sleep 100
        Click "Down"
        MouseMove 182, 310
        Sleep 100
        Click "Up"
        MouseMove 300, 307
        Sleep 100
        Send "{Backspace}"
        Sleep 100
        Send "{Text}TGDA"
        loop roomQty.Value {
            BlockInput true
            MouseMove 425, 537
            Sleep 200
            Send "!r"
            MouseMove 469, 512
            Sleep 3000
            Click
            Sleep 700
            Send "!t"
            Sleep 100
            Send "!s"
            Sleep 100
            Send "!m"
            Sleep 100
            Sleep 3000
            Send "{Esc}"
            Sleep 300
            Send "{Text}1"
            Sleep 100
            MouseMove 487, 402
            Sleep 100
            Click "Down"
            MouseMove 518, 403
            Sleep 200
            Click "Up"
            MouseMove 516, 443
            Sleep 200
            Send "{Text}0"
            Sleep 200
            Send "!o"
            Sleep 1000
            Send "!r"
            Sleep 200
            MouseMove 942, 598
            Sleep 2500
            Click
            Sleep 800
            Send "!d"
            Sleep 1000
            Send "{Left}"
            Sleep 150
            Send "{Space}"
            Sleep 500
            Send "!o"
            Sleep 500
            Send "!c"
            Sleep 300
            MouseMove 321, 507
            Sleep 1500
            Click "Down"
            MouseMove 186, 509
            Sleep 200
            Click "Up"
            Sleep 200
            Send "{Text}NRR"
            Sleep 100
            Send "{Tab}"
            Sleep 500
            loop 5 {
                Send "{Esc}"
                Sleep 200
            }
            Send "!o"
            Sleep 800
            Send "!o"
            Sleep 3000
            Send "!c"
            Sleep 800
            Send "!c"
            Sleep 300
            Send "!c"
            Sleep 5000
            BlockInput false
        }
        Sleep 1000
        MsgBox("已完成DNM & Share共 " . roomQty.Value . " 房，请核对有否错漏。", this.popupTitle, "4096")
    }
}