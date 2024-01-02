#Include "./utils.ahk"
#Include "../Lib/Classes/utils.ahk"
#Include "./FSR2 test script.ahk"

#SingleInstance Force
TraySetIcon A_ScriptDir . "\assets\FSRTray.ico"
version := "3.0.0"
popupTitle := "Fedex Scheduled Reservations " . version
config := A_ScriptDir . "\config.ini"
path := IniRead(config, "FSR", "schedulePath")
fsrDesc := "
	(
	快捷键及对应功能：

	Enter:     开始录入
	F12:       强制停止脚本

	)"

FSR := Gui("", popupTitle)
FSR.OnEvent("Close", (*) => quitApp())
FSR_UI := [
	FSR.AddText("", fsrDesc),
	FSR.AddEdit("vschdPath h25 w150 x20 y+10", path),
	FSR.AddButton("vselectFileBtn h25 w70 x+15", "选择文件"),
	FSR.AddButton("vopenFileBtn h25 w70 x+15", "打开文件"),
	FSR.AddText("x20 y+30 ", "2. 请指定提前一天留房时间点：`n（ETA在此时间点后的房间不会提前一天留房）"),
	FSR.AddDropDownList("vtoNextDropdown w150 x20 y+10 Choose2", ["09:00", "10:00", "11:00", "12:00", "13:00"]),
	FSR.AddText("x20 y+30 ", "3. 请选择需要录入预订的 Schedule 日期"),
	FSR.AddDateTime("vselectedDate x20 y+10"),
	FSR.AddButton("vshowQueryDateDetails x20 y+10", "显示Schedule"),
	FSR.AddText("x20 y+30 ", "※ 清除Schedule Excel占用：`n（此功能会强制关闭所有Excel 文件，请先确保已保存修改）"),
	FSR.AddButton("vkillExcelBtn h25 w70 y+15", "清除Excel"),
]
FSR.Show()

schdPath := interface.getCtrlByName("schdPath", FSR_UI)
selectFileBtn := interface.getCtrlByName("selectFileBtn", FSR_UI)
openFileBtn := interface.getCtrlByName("openFileBtn", FSR_UI)
toNextDropdown := interface.getCtrlByName("toNextDropdown", FSR_UI)
selectedDate := interface.getCtrlByName("selectedDate", FSR_UI)
showQueryDateDetails := interface.getCtrlByName("showQueryDateDetails", FSR_UI)
killExcelBtn := interface.getCtrlByName("killExcelBtn", FSR_UI)

schdPath.OnEvent("LoseFocus", savePath)
selectFileBtn.OnEvent("Click", getSchd)
openFileBtn.OnEvent("Click", (*) => Run(schdPath.Text))
toNextDropdown.OnEvent("Change", saveTime)
showQueryDateDetails.OnEvent("Click", (*) => showScheduleList(schdPath.Text))
killExcelBtn.OnEvent("Click", killExcel)

savePath(*) {
	IniWrite(schdPath.Text, config, "FSR", "schedulePath")
}

getSchd(*) {
	FSR.Opt("+OwnDialogs")
	selectFile := FileSelect(3, , "请选择 Schedule 文件")
	if (selectFile = "") {
		MsgBox("请选择文件")
		return
	}
	schdPath.Value := selectFile
	IniWrite(selectFile, config, "FSR", "schedulePath")
}

saveTime(*) {
	IniWrite(toNextDropdown.Text, config, "FSR", "resvOnDayTime")
}

killExcel(*) {
	if (ProcessExist("et.exe")) {
		ProcessClose "et.exe"
	} else if (ProcessExist("EXCEL.exe")) {
		ProcessClose "EXCEL.exe"
	} else {
		return
	}
}

showScheduleList(scheduledPath) {
	colTitles := [
		"Trip No.  ",
		"Qty",
		"Arr. Flight",
		"Arr. Date",
		"Arr. Time",
		"Stay Hours",
		"Dep. Date",
		"Dep. Time",
		"Dep. Flight",
	]
	Xl := ComObject("Excel.Application")
	schdBook := Xl.Workbooks.Open(scheduledPath)
	; retrieve schedule information
	flights := FedexScheduledReservations.getQueryDateFlights(schdBook, selectedDate.Value)
	Xl.Quit()

	SchdList := Gui("", Format("FedEx Schedule {1}", FormatTime(selectedDate.Value, "yyyy/MM/dd")))
	LV := SchdList.AddListView("Checked w600 h350", colTitles)
	; render list
	SchdDetails := []
	for flight in flights {
		SchdDetails.Push(
			LV.Add("Check",
				flight["tripNum"],
				flight["roomQty"],
				flight["flightIn1"] . flight["flightIn2"],
				flight["ibDate"],
				flight["ETA"],
				flight["stayHours"],
				flight["obDate"],
				flight["ETD"],
				flight["flightOut1"] . flight["flightOut2"]
			)
		)
	}

	SchdList.AddButton("y+10", "开始录入").OnEvent("Click", start)
	SchdList.OnEvent("Close", cleanReload)
	SchdList.Show()

	start(*) {
		flightKeys := [
			"tripNum",
			"roomQty",
			"flightIn",
			"ibDate",
			"ETA",
			"stayHours",
			"obDate",
			"ETD",
			"flightOut",
		]

		getCheckedRowNumbers(listViewCtrl, listData) {
			checkedRowNumbers := []
			loop listData.Length {
				curRow := listViewCtrl.GetNext(A_Index - 1, "Checked")
				try {
					if (curRow = prevRow || curRow = 0) {
						Continue
					}
				}
				checkedRowNumbers.Push(curRow)
				prevRow := curRow
			}
			return checkedRowNumbers
		}

		filterCheckedRowDataMap(listViewCtrl, mapKeys, checkedRows) {
			checkedRowsData := []
			for rowNumber in checkedRows {
				dataMap := Map()
				for key in mapKeys {
					dataMap[key] := listViewCtrl.GetText(rowNumber, A_Index)
				}
				checkedRowsData.Push(dataMap)
			}
			return checkedRowsData
		}

		checkedFlights := filterCheckedRowDataMap(LV, flightKeys, getCheckedRowNumbers(LV, flights))

		; selectedFlightList := []
		; store checked row numbers
		; checkedRows := []
		; loop flights.Length {
		; 	curRow := LV.GetNext(A_Index - 1, "Checked")
		; 	try {
		; 		if (curRow = prevRow || curRow = 0) {
		; 			Continue
		; 		}
		; 	}
		; 	checkedRows.Push(curRow)
		; 	prevRow := curRow
		; }

		; loop through every field in checked rows
		; loop checkedRows.Length {
		; 	flightMap := Map()
		; 	row := checkedRows[A_Index]
		; 	loop flightKeys.Length {
		; 		flightMap[flightKeys[A_Index]] := LV.GetText(row, A_Index)
		; 	}
		; 	selectedFlightList.Push(flightMap)
		; }
		FSR.Hide()
		SchdList.Hide()

		; FedexScheduledReservations.writeReservations(checkedFlights)
	}
}