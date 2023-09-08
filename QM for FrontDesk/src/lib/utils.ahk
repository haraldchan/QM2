cleanReload(){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    ; Key/Mouse state set default
    BlockInput false
	SetCapsLockState false
    CoordMode "Mouse", "Screen"
    ; Excel Quit
    ComObject("Excel.Application").Quit()
	Reload
}

quitApp() {
    quitConfirm := MsgBox("是否确定退出QM 2? ", "QM for FrontDesk 2.0.0", "OKCancel 4096")
    if (quitConfirm = "OK") {
        cleanReload()
        ExitApp
    } else {
        cleanReload()
    }
}