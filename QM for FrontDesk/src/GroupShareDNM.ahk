; reminder: Y-pos needs to minus 20
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

    selector := MsgBox(textMsg, "Q.Group Share&DNM", "YesNoCancel")
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
    roomQty := InputBox("`n请输入需要DNM的房间数量", "Q.Group Share&DNM", "h150")
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

; dnm share
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
    Send "TGDA"

    loop roomQty.Value {
        BlockInput true
        MouseMove 425, 537
        Sleep 200
        ; KeyDown "Alt", 1
        ; Sleep 75
        ; KeyDown "R", 1
        ; Sleep 80
        ; KeyUp "R", 1
        ; Sleep 6
        ; KeyUp "Alt", 1
        Send "!r"
        MouseMove 469, 512
        Sleep 3000
        Click
        Sleep 700
        ; KeyDown "Alt", 1
        ; Sleep 194
        ; KeyDown "T", 1
        ; Sleep 81
        ; KeyUp "T", 1
        ; Sleep 84
        ; KeyDown "S", 1
        ; Sleep 94
        ; KeyUp "S", 1
        ; Sleep 119
        ; KeyDown "M", 1
        ; Sleep 63
        ; KeyUp "M", 1
        ; Sleep 28
        ; KeyUp "Alt", 1
        Send "!t"
        Sleep 100
        Send "!s"
        Sleep 100
        Send "!m"
        Sleep 100
        Sleep 3000
        Send "{Esc}"
        Sleep 100
        Send "1"
        Sleep 100
        MouseMove 487, 402
        Sleep 100
        Click "Down"
        MouseMove 518, 403
        Sleep 200
        Click "Up"
        MouseMove 516, 443
        Sleep 200
        Send "0"
        Sleep 200
        ; KeyDown "Alt", 1
        ; Sleep 70
        ; KeyDown "O", 1
        ; Sleep 77
        ; KeyUp "O", 1
        ; Sleep 54
        ; KeyUp "Alt", 1
        Send "!o"
        Sleep 1000
        ; KeyDown "Alt", 1
        ; Sleep 23
        ; KeyDown "R", 1
        ; Sleep 81
        ; KeyUp "Alt", 1
        ; Sleep 3
        ; KeyUp "R", 1
        Send "!r"
        Sleep 200
        MouseMove 942, 598
        Sleep 2800
        Click
        Sleep 1500
        ; KeyDown "Alt", 1
        ; Sleep 40
        ; KeyDown "D", 1
        ; Sleep 77
        ; KeyUp "D", 1
        ; Sleep 15
        ; KeyUp "Alt", 1
        Send "!d"
        Sleep 1200
        Send "{Left}"
        Sleep 150
        ; KeyDown "Space", 1
        ; Sleep 42
        ; KeyUp "Space", 1
        Send "{Space}"
        Sleep 500
        ; KeyDown "Alt", 1
        ; Sleep 84
        ; KeyDown "O", 1
        ; Sleep 73
        ; KeyUp "O", 1
        ; Sleep 40
        ; KeyUp "Alt", 1
        Send "!o"
        Sleep 500
        ; KeyDown "Alt", 1
        ; Sleep 40
        ; KeyDown "C", 1
        ; Sleep 69
        ; KeyUp "Alt", 1
        ; Sleep 23
        ; KeyUp "C", 1
        Send "!c"
        Sleep 300
        MouseMove 321, 507
        Sleep 2000
        Click "Down"
        MouseMove 186, 509
        Sleep 200
        Click "Up"
        Sleep 200
        ; KeyDown "N", 1
        ; Sleep 85
        ; KeyUp "N", 1
        ; Sleep 40
        ; KeyDown "R", 1
        ; Sleep 72
        ; KeyUp "R", 1
        ; Sleep 53
        ; KeyDown "R", 1
        ; Sleep 90
        Send "{Text}NRR"
        Sleep 100
        ; KeyDown "Alt", 1
        ; Sleep 62
        ; KeyDown "O", 1
        ; Sleep 73
        ; KeyUp "O", 1
        ; Sleep 46
        ; KeyUp "Alt", 1
        ; Sleep 945
        Send "!o"
        Sleep 500
        ; KeyDown "Esc", 1
        ; Sleep 64
        ; KeyUp "Esc", 1
        ; Sleep 501
        ; KeyDown "Esc", 1
        ; Sleep 68
        ; KeyUp "Esc", 1
        ; Sleep 502
        ; KeyDown "Esc", 1
        ; Sleep 67
        ; KeyUp "Esc", 1
        ; Sleep 507
        ; KeyDown "Esc", 1
        ; Sleep 86
        ; KeyUp "Esc", 1
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        Send "{Esc}"
        Sleep 200
        ; KeyDown "Alt", 1
        ; Sleep 32
        ; KeyDown "O", 1
        ; Sleep 94
        ; KeyUp "O", 1
        ; Sleep 55
        ; KeyUp "Alt", 1
        ; Sleep 1618
        Send "!o"
        Sleep 1500
        ; KeyDown "Alt", 1
        ; Sleep 32
        ; KeyDown "C", 1
        ; Sleep 80
        ; KeyUp "Alt", 1
        ; Sleep 28
        ; KeyUp "C", 1
        ; Sleep 862
        Send "!c"
        Sleep 200
        ; KeyDown "Alt", 1
        ; Sleep 23
        ; KeyDown "C", 1
        ; Sleep 82
        ; KeyUp "Alt", 1
        ; Sleep 28
        ; KeyUp "C", 1
        Send "!c"
        Sleep 200
        BlockInput false
    }

}

; hotkeys
; F4:: GroupShareDNMMain()
; F12:: Reload  ; use 'Reload' for script reset
; ^F12:: ExitApp    ; use 'ExitApp' to kill script
