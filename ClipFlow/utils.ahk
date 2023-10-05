strToArr(str) {
    return StrSplit(str, ",")
}

arrToStr(arr) {
    str := ""
    loop arr.Length {
        str .= arr[A_Index] . ","
    }
    return str
}

cleanReload(quit := 0){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    if (WinExist("旅客信息")) {
        WinSetAlwaysOnTop false, "旅客信息"
    }
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    CoordMode "Mouse", "Screen"
    if (quit = "quit") {
        ExitApp
    }
    Reload
}

quitApp() {
    quitConfirm := MsgBox("是否确定退出QM 2? ", "QM for FrontDesk 2.1.0", "OKCancel 4096")
    quitConfirm = "OK" ? cleanReload("quit") : cleanReload()
}