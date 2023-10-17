class CityLedgerCo {
	static USE() {
		WinMaximize "ahk_class SunAwtFrame"
		WinActivate "ahk_class SunAwtFrame"
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		isBlue := "0x004080"
		PixelGetColor(600, 830) = isBlue
			? this.fullWin()
			: this.smallWin()
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}

	static fullWin() {
		BlockInput true
		MouseMove 862, 272
		Sleep 100
		Click
		Sleep 10
		Send "!o"
		Sleep 10
		Send "{Blind}{Text}CL"
		Sleep 10
		Send "!f"
		Sleep 10
		Send "!p"
		Sleep 10
		Send "{Esc}"
		Sleep 10
		MouseMove 894, 722
		Sleep 100
		BlockInput false
	}
	
	static smallWin() {
		BlockInput true
		MouseMove 700, 200
		Sleep 10
		Click
		Sleep 10
		Send "!o"
		Sleep 10
		Send "{Blind}{Text}CL"
		Sleep 10
		Send "!f"
		Sleep 10
		Send "!p"
		Sleep 10
		Send "{Esc}"
		Sleep 10
		BlockInput false
		Sleep 10
		WinRestore "ahk_class SunAwtFrame"
	}
}