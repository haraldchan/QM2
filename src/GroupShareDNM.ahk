; reminder: Y-pos needs to minus 20
GroupShareDnmMain() {
    ; bring OperaPMS window upfront and maximize
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"

    ; MsgBox options
    myTextYes := "是(Y) = 进行DNM和Share"
    myTextNo := "否(N) = 只进行DNM"
    myTextCancel := "取消 = 退出脚本"
    myTextMain := "`n将开始团Share及DNM"
    myText := Format("{1}`n{2}`n{3}`n`n{4}", myTextYes, myTextNo, myTextCancel, myTextMain)
    myTitle := "Q.Group Share&DNM"

    selector := MsgBox(myText, myTitle, "YesNoCancel")
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

    loop roomQty {
        BlockInput true
        MouseMove 430, 535
        Sleep 700
        Send "!r"
        Sleep 200
        MouseMove 474, 499
        Sleep 2500
        Click
        Sleep 593
        Send "!t"
        Sleep 200
        Send "!s"
        Sleep 200
        Send "!m"
        Sleep 200
        Send "{Esc}"
        Sleep 1000
        MouseMove 481, 383
        Sleep 100
        Click "Down"
        MouseMove 534, 393
        Sleep 130
        Click "Up"
        MouseMove 539, 374
        Sleep 200
        Send "0"
        MouseMove 596, 336
        Sleep 200
        Click
        MouseMove 654, 396
        Sleep 200
        Send "1"
        MouseMove 706, 576
        Sleep 560
        Click "Down"
        MouseMove 710, 576
        Sleep 64
        Click "Up"
        MouseMove 895, 603
        Sleep 450
        Click
        MouseMove 949, 578
        Sleep 2000
        Click
        MouseMove 898, 573
        Sleep 1000
        Click
        MouseMove 817, 549
        Sleep 750
        Send "{Enter}"
        MouseMove 861, 590
        Sleep 580
        Click
        MouseMove 317, 486
        Sleep 1400
        Click "Down"
        MouseMove 243, 492
        Sleep 180
        Click "Up"
        MouseMove 275, 490
        Sleep 550
        Send "{Text}NRR"
        Sleep 100
        Send "!o"
        Sleep 200
        MouseMove 638, 521
        Sleep 950
        Click
        MouseMove 660, 525
        Sleep 600
        Click
        Sleep 300
        Click
        MouseMove 637, 522
        Sleep 450
        Click
        MouseMove 659, 532
        Sleep 1800
        Send "!o"
        Sleep 1300
        Send "!c"
        Sleep 600
        Send "!c"
        Sleep 500
        BlockInput false
    }

}

; hotkeys
; F4:: GroupShareDNMMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
