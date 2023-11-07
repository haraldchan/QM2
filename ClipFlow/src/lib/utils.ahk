strToArr(str) {
    return StrSplit(str, ",")
}

arrToStr(arr) {
    str := ""
    loop arr.Length {
        str .= arr[A_Index] . ","
    }
    str := SubStr(str, 1, StrLen(str)-1)
    return str
}

filePrepend(textToInsert, fileToPrepend){
    textOrigin := FileRead(fileToPrepend)
    FileDelete fileToPrepend
    FileAppend textToInsert . textOrigin, fileToPrepend
}


getCtrlByName(vName, ctrlArray){
    loop ctrlArray.Length {
        if (vName = ctrlArray[A_Index].Name) {
            return ctrlArray[A_Index]
        }
    }
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

quitApp(*) {
    quitConfirm := MsgBox("是否确定退出ClipFlow? ", popupTitle, "OKCancel 4096")
    if (quitConfirm = "OK") {
        cleanReload("quit")
    } else {
        cleanReload()
    }
}