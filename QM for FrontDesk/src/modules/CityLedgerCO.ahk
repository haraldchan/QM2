class CityLedgerCo {
	static USE(initX := 0, initY := 0) {
		WinMaximize "ahk_class SunAwtFrame"
		WinActivate "ahk_class SunAwtFrame"
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		isBlue := "0x004080"
		PixelGetColor(600, 830) = isBlue
			? this.fullWin()
			: this.smallWin()
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}

	static fullWin(initX := 862, initY := 272) {
		BlockInput true
		MouseMove initX, initY ; 862, 272
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
		MouseMove initX - 517, initY - 3 ; 345, 269
		Sleep 100
		BlockInput false
	}

	static smallWin(initX := 700, initY := 200) {
		BlockInput true
		MouseMove initX, initY ; 700, 200
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