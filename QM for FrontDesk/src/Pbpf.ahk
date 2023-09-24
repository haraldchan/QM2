; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"

class PbPf {
	static description := "粘贴PayBy PayFor"

	static Main() {
		WinMaximize "ahk_class SunAwtFrame"
        WinActivate "ahk_class SunAwtFrame"
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		BlockInput true
		MouseMove 316, 699
		Sleep 300
		Send "!e"
		Sleep 200
		Send "^v"
		Sleep 150
		Send "!o"
		Sleep 100
		Send "!c"
		Sleep 100
		Send "!t"
		MouseMove 759, 266
		Sleep 200
		Click
		Send "!n"
		Sleep 200
		Send "{Text}OTH"
		MouseMove 517, 399
		Sleep 100
		Click
		MouseMove 479, 435
		Sleep 100
		Click
		MouseMove 689, 477
		Sleep 100
		Click "Down"
		MouseMove 697, 477
		Sleep 100
		Click "Up"
		Sleep 100
		Send "^v"
		Sleep 150
		Send "!o"
		Sleep 400
		Send "!c"
		Sleep 200
		Send "!c"
		Sleep 200
		BlockInput false
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}
}