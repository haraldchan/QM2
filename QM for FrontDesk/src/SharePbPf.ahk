; reminder: Y-pos needs to minus 20
; Main runs the main process
popupTitle := "INH Share & P/B P/F"
SharePbPfMain() {
	; bring OperaPMS window upfront and maximize
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"

	textMsg := "
	(
	是(Y) = InHouse Share
	否(N) = 粘贴P/B P/F信息
	取消 = 退出脚本

	INH Share：请在InHouse界面，窗口最大化情况下启动)
	粘贴P/B P/F：请在InHouse点开Comment窗口后启动
	)"
	selector := MsgBox(textMsg, popupTitle, "YesNoCancel")
	if (selector = "Yes") {
		share()
	} else if (selector = "No") {
		pbpf()
	} else {
		Reload
	}
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
}

share() {
	BlockInput true
	Sleep 100
	Send "!t"
	Sleep 200
	Send "!s"
	Sleep 200
	Send "!m"
	Sleep 700
	Send "{Esc}"
	Sleep 800
	Send "{Text}1"
	Sleep 364
	Send "{Tab}"
	Sleep 50
	Send "{Tab}"
	Sleep 50
	Send "{Tab}"
	Sleep 50
	Send "{Tab}"
	Sleep 50
	Send "{Text}0"
	Sleep 210
	Send "{Tab}"
	Sleep 50
	Send "{Tab}"
	Sleep 207
	Send "{Text}6"
	Sleep 1500
	Send "!o"
	Sleep 800
	Send "!r"
	Sleep 1000
	MouseMove 949, 579
	Sleep 2000
	Click
	Sleep 200
	Send "!d"
	MouseMove 611, 526
	Sleep 200
	Click
	Sleep 1500
	Send "!c"
	MouseMove 324, 487
	Sleep 1000
	Click "Down"
	MouseMove 212, 500
	Sleep 200
	Click "Up"
	Sleep 200
	Send "{Text}NRR"
	Sleep 200
	Send "!o"
	Sleep 1000
	Send "{Esc}"
	Sleep 500
	Send "{Esc}"
	Sleep 500
	Send "{Esc}"
	Sleep 4000
	Send "!i"
	Sleep 3000
	Send "{Esc}"
	Sleep 200
	Send "{Esc}"
	Sleep 1000
	Send "{Space}"
	Sleep 1500
	Send "!c"
	Sleep 300
	Send "!c"
	Sleep 300
	; share process end
	BlockInput false
}

pbpf() {
	BlockInput true
	; pbpf paste process start
	MouseMove 316, 679
	Sleep 300
	Send "!e"
	Sleep 200
	Send A_ClipBoard
	Sleep 150
	Send "!o"
	Sleep 100
	Send "!c"
	Sleep 100
	Send "!t"
	MouseMove 759, 246
	Sleep 200
	Click
	Send "!n"
	Sleep 200
	Send "{Text}OTH"
	MouseMove 517, 379
	Sleep 100
	Click
	MouseMove 479, 415
	Sleep 100
	Click
	MouseMove 689, 457
	Sleep 100
	Click "Down"
	MouseMove 697, 457
	Sleep 100
	Click "Up"
	Sleep 100
	Send A_ClipBoard
	Sleep 150
	Send "!o"
	Sleep 400
	Send "!c"
	Sleep 200
	Send "!c"
	Sleep 200
	BlockInput false
}

; hotkeys
; F9:: SharePbPfMain()
; F12:: Reload	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script
