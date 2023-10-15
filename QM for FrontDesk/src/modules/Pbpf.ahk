; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"

class PbPf {
	static description := "生成 PayBy PayFor 信息"
	static popupTitle := "PayBy PayFor"

	static Main() {
		static relations := []
		; GUI
		Main := Gui(, this.popupTitle)
		Main.AddGroupBox("w180 h190", "P/F房(支付人)")
		Main.AddText("xp+10 yp+25", "房号       ")
		pfRoom := Main.AddEdit("x+10 w80", "")
		Main.AddText("x20 y+10", "姓名/确认号")
		pfName := Main.AddEdit("x+10 w80", "")
		Main.AddText("x20 y+10", "Party号    ")
		party := Main.AddEdit("x+10 w80", "")
		Main.AddText("x20 y+10", "Total房数  ")
		roomQty := Main.AddEdit("x+10 w80", "")
		pfCopy := Main.AddButton("x20 y+10 h30 w160", "复制Pay For信息")

		Main.AddGroupBox("x+20 y8 w180 h190", "P/B房(被支付人)")
		Main.AddText("xp+10 yp+25", "房号       ")
		pbRoom := Main.AddEdit("x+10 w80", "")
		Main.AddText("xp-75 y+10", "姓名/确认号")
		pbName := Main.AddEdit("x+10 w80", "")
		pbCopy := Main.AddButton("xp-75 y+70 h30 w160", "复制Pay By信息")

		Main.AddButton("w370 x10 y+20 h35", "开始粘贴").OnEvent("Click", this.Run)

		pbpfCtrls := [pfRoom, pfName, party, roomQty, pbRoom, pbName]
		loop pbpfCtrls.Length {
			pbpfCtrls[A_Index].OnEvent("Change", update)
		}
		pfCopy.OnEvent("Click", getPayFor)
		pbCopy.OnEvent("Click", getPayBy)

		Main.Show()
		WinSetAlwaysOnTop true, this.popupTitle
		Main.OnEvent("Close", close)

		; callbacks
		close(*) {
			Main.Destroy()
		}

		update(*) {
			relations := [
				pfRoom.Value,
				pfName.Value,
				party.Value,
				roomQty.Value,
				pbRoom.Value,
				pbName.Value,
			]
		}

		getPayFor(*) {
			if (relations.Length = 0) {
				return
			}
			nameConf := IsNumber(relations[6]) ? "#" . relations[6] : relations[6]
			if (relations[3] = "" || relations[4] = "") {
				; 2-room party
				A_Clipboard := Format("P/F Rm{1}{2}  ", relations[5], nameConf)
			} else {
				; 3 or more room party
				A_Clipboard := Format("P/F Party#{1}, total {2}-rooms  ", relations[3], relations[4])

			}
			MsgBox(A_Clipboard, "已复制信息", "4096 T1")
		}

		getPayBy(*) {
			if (relations.Length = 0) {
				return
			}
			nameConf := IsNumber(relations[2]) ? "#" . relations[2] : relations[2]
			A_Clipboard := Format("P/B Rm{1}{2}  ", relations[1], nameConf)
			MsgBox(A_Clipboard, "已复制信息", "4096 T1")
		}
	}

	static Run(*) {
		commentPos := (A_OSVersion = "6.1.7601")
			? "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\QM for FrontDesk\src\assets\commentWin7.PNG"
			: "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\QM for FrontDesk\src\assets\comment.PNG"

		WinMaximize "ahk_class SunAwtFrame"
		WinActivate "ahk_class SunAwtFrame"
		; WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		BlockInput true
		CoordMode "Pixel", "Screen"
		CoordMode "Mouse", "Screen"
		if (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenHeight, commentPos)){
			anchorX := FoundX
			anchorY := FoundY
			MouseMove anchorX + 1, anchorY + 1
			Click
		} else {
			BlockInput false
			manualClick := MsgBox("请先点击打开Comment。", Pbpf.popupTitle, "OKCancel 4096")
			if (manualClick = "Cancel") {
				return
			}
			BlockInput true
		}
		Sleep 100
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
		; WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}
}