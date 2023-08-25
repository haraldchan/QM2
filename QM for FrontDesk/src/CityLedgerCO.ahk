; reminder: Y-pos needs to minus 20
; Main runs the main process

; findColor location: X600, Y830
; # = 0x
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
	MouseMove 862, 252
	Sleep 100
	Click
	Sleep 100
	Send "!o"
	Sleep 1500
	Send "{Blind}{Text}CL"
	Sleep 150
	Send "!f"
	Sleep 100
	Send "!p"
	Sleep 200
	MouseMove 669, 515
	Sleep 1000
	Click
	Sleep 300
	MouseMove 894, 702
	Sleep 100
	BlockInput false
}

smallWin() {
	BlockInput true
	MouseMove 700, 180
	Sleep 100
	Click
	Sleep 100
	; MouseMove 660, 605
	; Sleep 50
	; Click
	Send "!o"
	Sleep 1500
	Send "{Blind}{Text}CL"
	Sleep 150
	Send "!f"
	Sleep 100
	Send "!p"
	Sleep 200
	MouseMove 662, 520
	Sleep 1000
	Click
	Sleep 300
	BlockInput false
	Sleep 100
	WinRestore "ahk_class SunAwtFrame"
}

; hotkeys
; ^o:: CityLedgerCOMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
