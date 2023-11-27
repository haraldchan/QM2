; #Include "%A_ScriptDir%\src\utils.ahk"
#Include "../src/utils.ahk"

class FedexAmendmentHandler {
	static path := A_ScriptDir . "\src\FedExPasteTemplate.xls"

	static USE() {
		amendmentInfoMap := this.getAmendmentInfo()
		staffNames := ""
		loop amendmentInfoMap["staffInfo"].Length {
			staffNames .= amendmentInfoMap["staffInfo"][A_Index] . "; "
		}

		AM := Gui(, "Fedex Amendment 信息")
		ui := [
			LV := AM.AddListView("r11 w500 ReadOnly", ["Amendment 信息项目", "详情"]),
			LV.Add(, "Amendment 类型", amendmentInfoMap["resvType"]),
			LV.Add(, "房间数量", amendmentInfoMap["roomQty"]),
			LV.Add(, "Trip No.", amendmentInfoMap["tripNum"]),
			LV.Add(, "抵达航班", amendmentInfoMap["flightIn"]),
			LV.Add(, "抵达日期", amendmentInfoMap["ibDate"]),
			LV.Add(, "抵达时间", amendmentInfoMap["ETA"]),
			LV.Add(, "离开航班", amendmentInfoMap["flightOut"]),
			LV.Add(, "离开日期", amendmentInfoMap["obDate"]),
			LV.Add(, "离开时间", amendmentInfoMap["ETD"]),
			LV.Add(, "在住时长", amendmentInfoMap["stayHours"]),
			LV.Add(, "计费天数", amendmentInfoMap["daysActual"]),
			LV.Add(, "机组人员姓名", staffNames),
		]
		LV.ModifyCol(1, 100)
		LV.ModifyCol(2, 50)

		; add ctrl conditionally(depends on resvType)
		if (amendmentInfoMap["resvType"] = "ADD") {
			ui.Push([
				AM.AddButton("vstartBtn w80 h30", "开始新增")
			])
		} else {
			ui.Push([
				AM.AddText("x20 h30", "请只输入要修改的订单确认号（如果有多个，请用空格分开）"),
				AM.AddEdit("vconfNumInput x20 y+10 h30 w80", ""),
				AM.AddButton("vstartBtn x+10 w50 h30", "开始修改")
			])
		}
		AM.Show()

		;get ctrls
		startBtn := getCtrlByName("startBtn", ui)
		confNumInput := getCtrlByName("confNum", ui)

		;add event
		startBtn.OnEvent("Click", start)

		;callbacks
		start(*) {
			; handle CHANGE confirmation numbers
			if (amendmentInfoMap["resvType"] = "CHANGE") {
				if (InStr(confNumInput.Text, " ")) {
					; type array
					confNum := StrSplit(Trim(confNumInput.Text, " "), " ")
					extendedInfoMap := this.addPmsRequiredInfo(amendmentInfoMap, confNum)

				} else {
					; type string
					confNum := Trim(confNumInput.Text, " ")
					extendedInfoMap := this.addPmsRequiredInfo(amendmentInfoMap)
				}
			}

			loop amendmentInfoMap["roomQty"] {
				amendmentInfoMap["resvType"] = "CHANGE"
				? this.handleChange(extendedInfoMap)
				: this.handleAdd(extendedInfoMap)
			}
		}
	}

	static getAmendmentInfo() {
		amendmentMap := Map()

		Xl := ComObject("Excel.Application")
		Xlbook := Xl.Workbooks.Open(this.path)
		amendment := Xlbook.Worksheets("Sheet1")
		lastRow := amendment.Cells(Xlbook.ActiveSheet.Rows.Count, "A").End(-4162).Row

		amendmentMap["resvType"] := StrSplit(amendment.Cells(9, 1).Text, " ")[1]
		readStartRow := (amendmentMap["resvType"] = "ADD") ? 23 : 28

		amendmentMap["trackingNum"] := SubStr(amendment.Cells(7, 1).Text, 14)
		amendmentMap["tripNum"] := StrSplit(amendment.Cells(8, 1).Text, " ")[2] . "/" . StrSplit(amendment.Cells(8, 1).Text, " ")[4]

		amendmentMap["roomQty"] := StrSplit(amendment.Cells(readStartRow, 1).Text, " ")[2]

		inBoundInfo := StrSplit(amendment.Cells(readStartRow + 1, 1).Text, " ")
		amendmentMap["flightIn"] := inBoundInfo[3] . inBoundInfo[4]
		amendmentMap["ibDate"] := getReformatDate(inBoundInfo[7])
		amendmentMap["ETA"] := get24Hr(inBoundInfo[8], inBoundInfo[9])

		outBoundInfo := StrSplit(amendment.Cells(readStartRow + 2, 1).Text, " ")
		amendmentMap["flightOut"] := outBoundInfo[3] . outBoundInfo[4]
		amendmentMap["obDate"] := getReformatDate(outBoundInfo[7])
		amendmentMap["ETD"] := get24Hr(outBoundInfo[8], outBoundInfo[9])

		amendmentMap["stayHours"] := getStayHours(amendmentMap["ibDate"], amendmentMap["ETA"], amendmentMap["obDate"], amendmentMap["ETD"])
		amendmentMap["daysActual"] := getDaysActual(amendmentMap["stayHours"])

		staffInfoText := amendment.Cells(lastRow - 3, 1).Text
		amendmentMap["staffInfo"] := getStaffNames(staffInfoText)

		Xl.Quit()
		return amendmentMap
	}

	static addPmsRequiredInfo(amendmentInfo, confirmations := 0) {
		pmsCiDate := (StrSplit(amendmentInfo["ETA"], ":")[1]) < 10
			? DateAdd(amendmentInfo["ibDate"], -1, "days")
			: amendmentInfo["ibDate"]
		pmsCoDate := amendmentInfo["obDate"]
		amendmentInfo["pmsNts"] := DateDiff(pmsCoDate, pmsCiDate, "days")
		if (amendmentInfo["resvType"] = "CHANGE") {
			amendmentInfo["confNum"] := IsNumber(confirmations) ? confirmations : StrSplit(confirmations, " ")
			amendmentInfo["comment"] := Format("Changed to {1}={2} days. Actual Stay: {3}-{4}`n", amendmentInfo["stayHours"], amendmentInfo["daysActual"], amendmentInfo["ibDate"], amendmentInfo["obDate"])
		} else {
			amendmentInfo["comment"] := Format("RM INCL 1BBF TO CO,Hours@Hotel:{1}={2}days, ActualStay: {3}-{4}", amendmentInfo["stayHours"], amendmentInfo["daysActual"], amendmentInfo["ibDate"], amendmentInfo["obDate"])
		}
		; reformat to match pms date format
		amendmentInfo["amCiDate"] := FormatTime(amendmentInfo["ibDate"], "MMddyyyy")
		amendmentInfo["amCoDate"] := FormatTime(amendmentInfo["obDate"], "MMddyyyy")
		amendmentInfo["pmsCiDate"] := FormatTime(pmsCiDate, "MMddyyyy")
		amendmentInfo["pmsCoDate"] := FormatTime(pmsCoDate, "MMddyyyy")
	}

	static handleChange(extendedInfoMap) {
		loop extendedInfoMap["roomQty"] {
			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
			; TODO: find the reservation needs to be update through "Update Reservation"
			
		}
	}

	static handleAdd(extendedInfoMap) {
		loop extendedInfoMap["roomQty"] {
			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
			; { opening a booking
			; TODO: adding a booking from PM
			; }
			Sleep 100
			; { opening Profile, fill in inbound flight info & trip no.
			MouseMove 467, 201
			Sleep 300
			Click
			Sleep 1000
			MouseMove 442, 264
			Sleep 300
			Click "Down"
			MouseMove 214, 269
			Sleep 300
			Click "Up"
			Sleep 300
			Send "{Backspace}"
			Sleep 300
			; entry: last name
			Send Format("{Text}{1}", StrSplit(extendedInfoMap["staffInfo"][A_Index], " ")[2])
			Sleep 300
			; TODO: entry first name action
			MouseMove 594, 394
			Sleep 150
			Send "!c"
			Sleep 300
			MouseMove 576, 521
			Sleep 150
			Click
			Sleep 300
			MouseMove 798, 480
			Sleep 500
			Click
			Sleep 500
			; }
			Sleep 2000
			; { fill in check-in, check-out fields
			MouseMove 332, 336
			Sleep 1000
			Click "Down"
			MouseMove 178, 340
			Sleep 300
			Click "Up"
			MouseMove 172, 340
			Sleep 300
			Send Format("{Text}{1}", extendedInfoMap["pmsCiDate"])
			Sleep 100
			MouseMove 325, 378
			Sleep 300
			Click
			Sleep 300
			MouseMove 661, 523
			Sleep 300
			Click
			MouseMove 636, 523
			Sleep 300
			Click
			MouseMove 635, 523
			Sleep 300
			Click
			Sleep 300
			Click
			Sleep 300
			MouseMove 335, 385
			Sleep 300
			Click "Down"
			MouseMove 182, 389
			Sleep 300
			Click "Up"
			MouseMove 207, 395
			Sleep 300
			Send Format("{Text}{1}", extendedInfoMap["pmsCoDate"])
			Sleep 300
			; }
			Sleep 2000
			; { fill in ETA & ETD
			MouseMove 294, 577
			Sleep 200
			Click
			Sleep 200
			Send "{Enter}"
			Sleep 200
			Send "{Enter}"
			Sleep 200
			Send "{Enter}"
			Sleep 200
			MouseMove 320, 577
			Sleep 200
			Click "Down"
			MouseMove 200, 577
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", extendedInfoMap["ETA"])
			Sleep 200
			MouseMove 499, 577
			Sleep 200
			Click "Down"
			MouseMove 330, 574
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", extendedInfoMap["ETD"])
			Sleep 200
			; }
			Sleep 200
			; { fill in comments, Inbound & trip no.(TA REC log field)
			MouseMove 622, 576
			Sleep 200
			Click "Down"
			MouseMove 1140, 585
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", extendedInfoMap["comment"])
			Sleep 200
			MouseMove 839, 535
			Sleep 100
			Click "Down"
			MouseMove 1107, 543
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}  {2}", extendedInfoMap["flightIn"], extendedInfoMap["tripNum"])
			Sleep 2000
			; }
			Sleep 100
			; { fill in original shcedule info in "More Fields"
			MouseMove 236, 313
			Sleep 100
			Click
			Sleep 100
			MouseMove 686, 439
			Sleep 200
			Click "Down"
			MouseMove 478, 439
			Sleep 100
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}", extendedInfoMap["flightIn"])
			Sleep 100
			MouseMove 672, 483
			Sleep 281
			Click "Down"
			MouseMove 523, 483
			Sleep 358
			Click "Up"
			Send Format("{Text}{1}", extendedInfoMap["amCiDate"])
			Sleep 100
			MouseMove 685, 506
			Sleep 100
			Click "Down"
			MouseMove 422, 501
			Sleep 100
			Click "Up"
			Send Format("{Text}{1}", extendedInfoMap["ETA"])
			Sleep 100
			MouseMove 922, 441
			Sleep 100
			Click "Down"
			MouseMove 704, 439
			Sleep 100
			Click "Up"
			Send Format("{Text}{1}", extendedInfoMap["flightOut"])
			Sleep 100
			MouseMove 917, 483
			Sleep 100
			Click "Down"
			MouseMove 637, 487
			Sleep 100
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}", extendedInfoMap["amCoDate"])
			Sleep 100
			MouseMove 922, 504
			Sleep 100
			Click "Down"
			MouseMove 640, 503
			Sleep 100
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}", extendedInfoMap["ETD"])
			Sleep 100
			MouseMove 841, 660
			Sleep 100
			Click
			Sleep 2000
			; }
			Sleep 100
			; { open "Daily Details", fill in room rates
			MouseMove 372, 504
			Sleep 300
			Click
			Sleep 300
			Send "!d"
			Sleep 300
			loop extendedInfoMap["daysActual"] {
				Send "{Text}1265"
				Send "{Down}"
				Sleep 200
			}
			if (extendedInfoMap["daysActual"] != extendedInfoMap["pmsNts"]) {
				Send "0"
				Sleep 200
				MouseMove 618, 485
				Sleep 200
				Send "!e"
				Sleep 200
				loop 4 {
					Send "{Tab}"
					Sleep 100
				}
				Send "{Text}NRR"
				Sleep 100
				MouseMove 418, 377
				Sleep 100
				Send "!o"
				Sleep 200
				loop 3 {
					Send "{Escape}"
					Sleep 200
				}
				Send "!o"
				Sleep 1500
				Send "{Space}"
				Sleep 200
			}
			; TODO: MIND THE MSGBOXES!!!!
			MouseMove 728, 548
			Sleep 300
			Send "!o"
			Sleep 300
			MouseMove 542, 453
			Sleep 300
			Click
			MouseMove 644, 523
			Sleep 300
			Click
			Sleep 300
			; }
			Sleep 2000
			; { close, save, down to the next one

			; }

		}
	}
}