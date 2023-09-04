; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"
; scoped vals
GroupShareDnmConfig := {
    popupTitle: "Group Share & DoNotMove"
}

GroupShareDnmMain() {
    ; bring OperaPMS window upfront and maximize
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    textMsg := "
    (
    是(Y) = 进行DNM和Share
    否(N) = 只进行DNM
    取消 = 退出脚本

    将开始团Share及DNM
    )"

    selector := MsgBox(textMsg, GroupShareDnmConfig.popupTitle, "YesNoCancel")
    if (selector = "Yes") {
        dnmShare()
    } else if (selector = "No") {
        dnm()
    } else {
        cleanReload()
    }
}

dnm() {
    roomQty := InputBox("`n请输入需要DNM的房间数量", GroupShareDnmConfig.popupTitle, "h150")
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
    MsgBox("已完成批量DoNotMove，合共" . roomQty . "房。", GroupShareDnmConfig.popupTitle)
}

dnmShare() {
    msgText := "
    (
    将开始批量团Share及DoNotMove操作。

    运行前请先确认：
    1.Opera窗口已最大化。
    2.界面在RoomAssign。
    3.以Name筛选团房（如使用BlockCode将会出错）
    )"
    confirmMsg := MsgBox(msgText, GroupShareDnmConfig.popupTitle, "OKCancel 4096")
    If (confirmMsg = "Cancel") {
        cleanReload()
    } else {
        roomQty := InputBox("`n请输入需要Share + DNM的房间数量", GroupShareDnmConfig.popupTitle, "h150")
        if (roomQty.Result = "Cancel") {
            cleanReload()
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
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "!o"
        Sleep 800
        Send "!o"
        Sleep 3000
        Send "!c"
        Sleep 800
        Send "!c"
        Sleep 300
        Send "!c"
        Sleep 2000
        BlockInput false
    }
    MsgBox("已完成DNM & Share，请核对有否错漏。", GroupShareDnmConfig.popupTitle, "4096")
}

; hotkeys
; F4:: GroupShareDNMMain()
; F12:: Reload  ; use 'Reload' for script reset
; ^F12:: ExitApp    ; use 'ExitApp' to kill script