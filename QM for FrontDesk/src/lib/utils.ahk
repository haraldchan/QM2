cleanReload(quit := 0){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        try {
            WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
        } catch {
            MsgBox("
            (
            当前为非管理员状态，将无法正确运行。
            请在QM 2 图标上右键点击，选择“以管理员身份运行”。    
            )", "QM for FrontDesk 2.1.0")
        }
    }
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    CoordMode "Mouse", "Screen"
    ; Excel Quit
    ComObject("Excel.Application").Quit()
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