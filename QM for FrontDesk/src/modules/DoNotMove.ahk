; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"

class DoNotMove {
    static description := "预到房间批量 DoNotMove"
    static popupTitle := "Do Not Move(batch)"

    static USE() {
        WinMaximize "ahk_class SunAwtFrame"
        WinActivate "ahk_class SunAwtFrame"
        Sleep 500
        roomQty := InputBox("`n请输入需要DNM的房间数量", this.popupTitle, "h150")
        if (roomQty.Result = "Cancel") {
            cleanReload()
        }
        BlockInput true
        loop roomQty.Value {
            MouseMove 696, 614
            Sleep 200
            Send "!r"
            Sleep 100
            MouseMove 579, 527
            Sleep 2000
            Click
            Sleep 400
            Click
            Sleep 300
            Click
            MouseMove 473, 514
            Sleep 500
            Click
            Sleep 700
            Send "!o"
            Sleep 1500
        }
        BlockInput false
        Sleep 1000
        MsgBox("已完成批量DoNotMove，合共" . roomQty.Value . "房。", this.popupTitle)
    }
}