; reminder: Y-pos needs to minus 20
popupTitle := "Group Share & DoNotMove"

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

    selector := MsgBox(textMsg, popupTitle, "YesNoCancel")
    if (selector = "Yes") {
        dnmShare()
    } else if (selector = "No") {
        dnm()
    } else {
        Reload
    }
}

; dnm only
dnm() {
    roomQty := InputBox("`n请输入需要DNM的房间数量", popupTitle, "h150")
    if (roomQty.Result = "Cancel") {
        Reload
    }
    BlockInput true
    loop roomQty.Value {
        MouseMove 696, 594
        Sleep 200
        Send "!r"
        Sleep 100
        MouseMove 579, 507
        Sleep 2000
        Click
        Sleep 400
        Click
        Sleep 300
        Click
        MouseMove 473, 494
        Sleep 500
        Click
        Sleep 700
        Send "!o"
        Sleep 1500
    }
    BlockInput false
    MsgBox(Format("已完成批量DoNotMove，合共 {1} 房。", roomQty.Value), "Q.Group Share&DNM")
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
    confirmMsg := MsgBox(msgText, "Q.Group Share&DNM", "OKCancel 4096")
    If (confirmMsg = "Cancel") {
        Reload
    } else {
        roomQty := InputBox("`n请输入需要Share + DNM的房间数量", "Q.Group Share&DNM", "h150")
        if (roomQty.Result = "Cancel") {
            Reload
        }
    }
    MouseMove 340, 291
    Sleep 100
    Click "Down"
    MouseMove 182, 290
    Sleep 100
    Click "Up"
    MouseMove 300, 287
    Sleep 100
    Send "{Backspace}"
    Sleep 100
    Send "{Text}TGDA"
    loop roomQty.Value {
        BlockInput true
        MouseMove 425, 517
        Sleep 200
        Send "!r"
        MouseMove 469, 492
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
        MouseMove 487, 382
        Sleep 100
        Click "Down"
        MouseMove 518, 383
        Sleep 200
        Click "Up"
        MouseMove 516, 423
        Sleep 200
        Send "{Text}0"
        Sleep 200
        Send "!o"
        Sleep 1000
        Send "!r"
        Sleep 200
        MouseMove 942, 578
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
        MouseMove 321, 487
        Sleep 1500
        Click "Down"
        MouseMove 186, 489
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
    MsgBox("已完成DNM & Share，请核对有否错漏。", popupTitle, "4096")
}

; hotkeys
; F4:: GroupShareDNMMain()
; F12:: Reload  ; use 'Reload' for script reset
; ^F12:: ExitApp    ; use 'ExitApp' to kill script
