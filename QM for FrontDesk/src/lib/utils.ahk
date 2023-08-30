cleanReload(){
    BlockInput false
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
	SetCapsLockState false
    CoordMode "Mouse", "Client"
	Reload
}