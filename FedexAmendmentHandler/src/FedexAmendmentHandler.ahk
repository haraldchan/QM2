; #Include "%A_ScriptDir%\src\utils.ahk"
#Include "../../src/utils.ahk"

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
			if (isObject(confNumInput)) {
				if (InStr(confNumInput.Text, " ")) {
					confNum := StrSplit(Trim(confNumInput.Text, " "), " ")
				} else {
					confNum := Trim(confNumInput.Text, " ")
				}
			}
			if (amendmentInfoMap["resvType"] = "CHANGE") {
				extendedInfoMap := this.addPmsRequiredInfo(amendmentInfoMap, confNum)
			} else {
				extendedInfoMap := this.addPmsRequiredInfo(amendmentInfoMap)
			}

			loop extendedInfoMap["roomQty"] {
				this.determineChangeAdd(extendedInfoMap)
			}
		}
	}

	static getAmendmentInfo() {
		amendmentMap := Map()
		; {
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

	static determineChangeAdd(extendedInfoMap) {
		if (extendedInfoMap["resvType"] = "CHANGE") {
			;TODO: find exist reservations, and modify it
			this.resvInfoEntry(extendedInfoMap)
		} else {
			;TODO: add-on new reservations from PM, then modify it
			this.resvInfoEntry(extendedInfoMap)
		}
	}

	static resvInfoEntry(extendedInfoMap) {
		;TODO: actions to fill in
	}


}