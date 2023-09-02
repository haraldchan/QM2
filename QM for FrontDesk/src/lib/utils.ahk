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