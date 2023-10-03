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
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    CoordMode "Mouse", "Screen"
    if (quit = "quit") {
        ExitApp
    }
    Reload
}