#Include "../Lib/Classes/utils.ahk"
#Include "./src/FSR_Utils.ahk"
#Include "./src/FedexScheduledResv.ahk"

#SingleInstance Force
CoordMode "Mouse", "Screen"
TraySetIcon A_ScriptDir . "\src\assets\FSRTray.ico"
version := "3.0.0"
popupTitle := "FSR " . version
config := A_ScriptDir . "\src\config.ini"
path := IniRead(config, "FSR", "schedulePath")
fsrDesc := "
	(
	快捷键及对应功能：

	Enter:     开始录入
	F12:       强制停止脚本

	)"

FSR := Gui("", popupTitle)
FSR.OnEvent("Close", (*) => quitApp())
FSR.OnEvent("DropFiles", dropSchd)
FSR_UI := [
	FSR.AddText("", fsrDesc),
	; FSR.AddText("x20 y+10", "1. 请选择Schedule 文件"),

	FSR.AddGroupBox("r4.5 w250 x10 y+10", "1. 选择 Schedule 文件"),
	FSR.AddText("x20 yp+20 ", "(直接拖动文件到此处即可)"),
	FSR.AddEdit("vschdPath h25 w150 x20 yp+20", path),
	FSR.AddButton("vselectFileBtn h25 w70 ", "选择文件"),
	FSR.AddButton("vopenFileBtn h25 w70 x+15", "打开文件"),

	FSR.AddGroupBox("r2.5 w250 x10 y+30", "2. 指定提前一天留房时间点"),
	FSR.AddText("x20 yp+20 ", "ETA在此时间点后的房间不会提前一天留房。"),
	FSR.AddDropDownList("vtoNextDropdown w150 x20 y+10 Choose2", ["09:00", "10:00", "11:00", "12:00", "13:00"]),

	FSR.AddGroupBox("r3 w250 x10 y+30", "3. 选择需要录入的航班日期"),
	FSR.AddText("x20 yp+20 ", "点击查看航班无误后即可开始"),
	FSR.AddDateTime("vselectedDate x20 y+10 w150"),
	FSR.AddButton("vshowQueryDateDetails Default x+15 h25 w70", "显示航班"),
]
FSR.Show()

schdPath := interface.getCtrlByName("schdPath", FSR_UI)
selectFileBtn := interface.getCtrlByName("selectFileBtn", FSR_UI)
openFileBtn := interface.getCtrlByName("openFileBtn", FSR_UI)
toNextDropdown := interface.getCtrlByName("toNextDropdown", FSR_UI)
selectedDate := interface.getCtrlByName("selectedDate", FSR_UI)
showQueryDateDetails := interface.getCtrlByName("showQueryDateDetails", FSR_UI)

schdPath.OnEvent("LoseFocus", savePath)
selectFileBtn.OnEvent("Click", getSchd)
openFileBtn.OnEvent("Click", (*) => Run(schdPath.Text))
toNextDropdown.OnEvent("Change", saveTime)
showQueryDateDetails.OnEvent("Click", (*) => showScheduleList(schdPath.Text))

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

dropSchd(GuiObj, GuiCtrlObj, FileArray, X, Y, *) {
    if (FileArray.Length > 1) {
        FileArray.RemoveAt[1]
    }
    schdPath.Value := FileArray[1]
    savePath()
}

saveTime(*) {
	IniWrite(toNextDropdown.Text, config, "FSR", "resvOnDayTime")
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
	SchdList.OnEvent("Close", cleanReload)
	LV := SchdList.AddListView("Checked w620 h350", colTitles)

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

	footer := [
		SchdList.AddCheckbox("vcheckAllBtn Checked h25 y+10", "全选"),
		SchdList.AddButton("vstartBtn Default x+10", "开始录入"),
	]

	SchdList.Show()

	global isCheckedAll := true
	checkAllBtn := interface.getCtrlByName("checkAllBtn", footer)
	startBtn := interface.getCtrlByName("startBtn", footer)

	startBtn.OnEvent("Click", start)
	checkAllBtn.OnEvent("Click", setCheckAll)
	LV.OnEvent("ItemCheck", checkIsCheckedAll)

	checkIsCheckedAll(*){
		checkedRows := interface.getCheckedRowNumbers(LV)
		checkAllBtn.Value := (checkedRows.Length = flights.Length)
		global isCheckedAll := checkAllBtn.Value
	}

	setCheckAll(*){
		global isCheckedAll := !isCheckedAll
		checkStatus := isCheckedAll = true
			? "Check"
			: "-Check"
			LV.Modify(0, checkStatus)
	}

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

		checkedFlights := interface.getCheckedRowDataMap(LV, flightKeys, interface.getCheckedRowNumbers(LV))

		FSR.Hide()
		SchdList.Hide()

		FedexScheduledReservations.writeReservations(checkedFlights)

		FSR.Show()
	}
}

; hotkey setups
F12:: cleanReload()
#Hotif WinActive("ahk_class AutoHotkeyGUI")
Esc:: quitApp()