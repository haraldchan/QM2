cleanReload(quit := 0){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    ; Quit App
    if (quit = "quit") {
        ExitApp
    }
    Reload
}

quitApp() {
    quitConfirm := MsgBox("是否确定退出QM 2? ", "QM for FrontDesk 2.1.0", "OKCancel 4096")
    if (quitConfirm = "OK") {
        cleanReload("quit")
    } else {
        cleanReload()
    }
}

getDaysActual(h, m) {
    if (h < 24) {
        return 1
    } else if (Mod(h, 24) = 0 && m = 0) {
        return Integer(h / 24)
    } else if (h >= 24 || m != 0) {
        return Integer(h / 24 + 1)
    }
}

getBringForwardTime(timeString) {
    switch timeString {
        case "09:00":
            return 9
        case "10:00":
            return 10
        case "11:00":
            return 11
        case "12:00":
            return 12
        case "13:00":
            return 13
        default:
            return 10
    }
}