; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../lib/utils.ahk"

class PbPf {
	static description := "粘贴PayBy PayFor"
	static popupTitle := "PayBy PayFor"

	static Main() {
		static relations := []
		WinSetAlwaysOnTop true, this.popupTitle
		; GUI
		Main := Gui(, this.popupTitle)
		Main.AddGroupBox("", "P/F房(支付人)")
		Main.AddText("xp+10", "房号")
		pfRoom := Main.AddEdit("x+10 w50", "")
		Main.AddText("xp+10", "姓名/确认号")
		pfName := Main.AddEdit("x+10 w50", "")
		Main.AddText("xp+10", "Party号")
		party := Main.AddEdit("x+10 w50", "")
		Main.AddText("xp+10", "Total房数")
		roomQty := Main.AddEdit("x+10 w50", "")
		pfCopy := Main.AddButton("xp+10 w50", "复制Pay For信息")

		Main.AddGroupBox("", "P/B房(被支付人)")
		Main.AddText("xp+10", "房号")
		pbRoom := Main.AddEdit("x+10 w50", "")
		Main.AddText("xp+10", "姓名/确认号")
		pbName := Main.AddEdit("x+10 w50", "")
		pbCopy := Main.AddButton("xp+10 w50", "复制Pay By信息")

		Main.AddButton("w200", "开始粘贴").OnEvent("Click", this.Run)

		pbpfCtrls := [pfRoom, pfName, party, roomQty, pbRoom, pbName]
		loop pbpfCtrls.Length {
			pbpfCtrls[A_Index].OnEvent("Change", update)
		}
		pfCopy.OnEvent("Click", getPayFor)
		pfCopy.OnEvent("Click", getPayBy)

		Main.Show()
		Main.OnEvent("Close", close)

		; callbacks
		close(*) {
			Main.Destroy()
		}

		update(*) {
			relations := [
				pfRoom.Text,
				pfName.Text,
				party.Text,
				roomQty.Text,
				pbRoom.Text,
				pbName.Text,
			]
		}

		getPayFor(*) {
			relations[6] := IsNumber(relations[6]) ? "#" . relations[6] : relations[6]
			if (relations[3] = "" || relations[4] = "") {
				; 2-room party
				A_Clipboard := Format("P/F Rm{1}{2}", relations[5], relations[6])
			} else {
				; 3 or more room party
				A_Clipboard := Format("P/F Party#{1}, total {2}-rooms", relations[3], relations[4])
			}
		}

		getPayBy(*) {
			relations[2] := IsNumber(relations[2]) ? "#" . relations[2] : relations[2]
			A_Clipboard := Format("P/B Rm{1}{2}", relations[1], relations[2])
		}
	}

	static Run(*) {
		WinMaximize "ahk_class SunAwtFrame"
		WinActivate "ahk_class SunAwtFrame"
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		BlockInput true
		ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, A_ScriptDir . "\src\assets\comment.PNG")
		anchorX := FoundX
		anchorY := FoundY
		MouseMove anchorX + 10, anchorY + 10
		Click
		Sleep 100
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