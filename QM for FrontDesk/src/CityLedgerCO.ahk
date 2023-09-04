CityLedgerCOMain() {
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	isBlue := "0x004080"
	if (PixelGetColor(600, 830) = isBlue) {
		fullWin()
	} else {
		smallWin()
	}
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
}

fullWin() {
	BlockInput true
	MouseMove 862, 272
	Sleep 100
	Click
	Sleep 100
	Send "!o"
	Sleep 1000
	Send "{Blind}{Text}CL"
	Sleep 150
	Send "!f"
	Sleep 100
	Send "!p"
	Sleep 200
	; MouseMove 669, 535
	; Sleep 1000
	; Click
	Send "{Esc}"
	Sleep 300
	MouseMove 894, 722
	Sleep 100
	BlockInput false
}

smallWin() {
	BlockInput true
	MouseMove 700, 200
	Sleep 100
	Click
	Sleep 100
	Send "!o"
	Sleep 1000
	Send "{Blind}{Text}CL"
	Sleep 150
	Send "!f"
	Sleep 100
	Send "!p"
	Sleep 200
	; MouseMove 662, 540
	; Sleep 1000
	; Click
	Send "{Esc}"
	Sleep 300
	BlockInput false
	Sleep 100
	WinRestore "ahk_class SunAwtFrame"
}

; hotkeys
; ^o:: CityLedgerCOMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
