; #Include "%A_ScriptDir%\Lib\reports.ahk"
; #Include "%A_ScriptDir%\Lib\utils.ahk"
#Include "../Lib/reports.ahk"
#Include "../Lib/utils.ahk"
; scoped vals
ReportMaster := {
	popupTitle: "Report Master",
}

overnightReportList := [
	comp,
	mgrFlash,
	hisFor15,
	hisForThisMonth,
	hisForNextMonth,
	vipArr,
	vipDep,
	arrAll,
	inhAll,
	depAll,
	creditLimit,
	bbf,
	rooms,
	ooo,
	groupRoom,
	groupInh,
	noShow,
	cancel,
	vipInh
]

overnightReportNames := [
	"1  - Guest INH Complimentary",
	"2  - NA02-Manager Flash",
	"3  - RS05-（前后15天）History & Forecast",
	"4  - RS05-（FO当月）History & Forecast",
	"5  - RS05-（FO次月）History & Forecast",
	"6  - FO01-VIP Arrival (VIP Arr)",
	"7  - FO03-VIP DEP",
	"8  - FO01-Arrival Detailed",
	"9  - FO02-Guests INH by Room",
	"10 - FO03-Departures",
	"11 - FO11-Credit Limit",
	"12 - FO13-Package Forecast（仅早餐）",
	"13 - Rooms-housekeepingstatus",
	"14 - HK03-OOO",
	"15 - Group Rooming List",
	"16 - Group In House",
	"17 - FO08-No Show",
	"18 - Cancellations"
	"19 - Guest In House w/o Due Out(VIP INH) ",
]

ReportMasterMain() {
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"
	reportMsg := "
	(
	请输入对应的报表编号（默认为夜审后操作）。
	（报表将保存至 开始菜单-文档）。

	夜班报表：
	1  - Guest INH Complimentary
	2  - NA02-Manager Flash
	3  - RS05-（前后15天）History & Forecast
	4  - RS05-（FO当月）History & Forecast
	5  - RS05-（FO次月）History & Forecast
	6  - FO01-VIP Arrival (VIP Arr)
	7  - FO03-VIP DEP
	8  - FO01-Arrival Detailed
	9  - FO02-Guests INH by Room
	10 - FO03-Departures
	11 - FO11-Credit Limit
	12 - FO13-Package Forecast（仅早餐）
	13 - Rooms-housekeepingstatus
	14 - HK03-OOO
	15 - Group Rooming List
	16 - Group In House
	17 - FO08-No Show
	18 - Cancellations
	19 - Guest In House w/o Due Out 
	666 - 保存以上所有报表（执行时间约8分钟，期间请勿操作电脑）

	其他：
	garr - 保存当天团单
	sp - 保存当天水果5（Excel格式）
	)"
	reportSelector := InputBox(reportMsg, ReportMaster.popupTitle, "h600", "666")
	if (reportSelector.Result = "Cancel") {
		cleanReload()
	}
	try { ; input value is number
		if (reportSelector.Value < 20 && reportSelector.Value > 0) {
			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
			reportName := overnightReportNames[reportSelector.Value]
			overnightReportList[reportSelector.Value]()
			Sleep 1000
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			openMyDocs(reportName)
		} else if (reportSelector.Value = 666) {
			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
			loop overnightReportList.Length {
				if (A_Index = 5) {
					if ((dateLast() - A_DD) > 5) {
						continue
					}
				}
				overnightReportList[A_Index]()
				Sleep 2000
			}
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			openMyDocs("夜班报表")
		} else {
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			MsgBox("请选择表中的指令", ReportMaster.popupTitle)
		}
	}
	catch { ; input value is string
		if (StrLower(reportSelector.Value) = "garr") {
			reportName := "当天 Arrival 团单"
			onDayGroup := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{2}\{2}{3}\{1}Group ARR&DEP.xlsx", today, A_Year, A_MM)
			blockInfo := getBlockInfo(onDayGroup)
			blockInfoText := ""
			for blockName, blockCode in blockInfo {
				blockInfoText .= Format("{1}：{2}`n", blockName, blockCode)
			}
			RmListSaver := MsgBox(Format("
					(	
					请确认当天Arrival 团队信息：

					{1}

					是(Y)：自动保存上述团队团单
					否(N)：手动录入block code保存团单
					取消：退出脚本
					)", blockInfoText), ReportMaster.popupTitle, "YesNoCancel")
			if (RmListSaver = "Yes") {
				groupArrAuto(blockinfo)
			} else if (RmListSaver = "No") {
				groupArr()
			} else {
				cleanReload()
			}
			Sleep 300
			openMyDocs(reportName)
		} else if (StrLower(reportSelector.Value) = "sp") {
			reportName := today . "水果5"
			special(today)
			Sleep 300
			openMyDocs(reportName)
		} else {
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			MsgBox("请选择表中的指令", ReportMaster.popupTitle)
		}
	}
}

openMyDocs(reportName) {
	WinSetAlwaysOnTop false
	myText := "已保存报表：" . reportName . "`n`n是否打开所在文件夹? "
	openFolder := MsgBox(myText, ReportMaster.popupTitle, "OKCancel")
	if (openFolder = "OK") {
		Run A_MyDocuments
	} else {
		cleanReload()
	}
}

dateLast() {
	preAuditDate := DateAdd(A_Now, -1, "Days")
	preAuditMonth := FormatTime(preAuditDate, "MM")
	preAuditYear := FormatTime(preAuditDate, "yyyy")
	nextMonth := (preAuditMonth = 12) ? 1 : preAuditMonth + 1
	if (nextMonth < 10) {
		nextMonth := "0" . nextMonth
	}
	printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
	firstDayOfNextMonth := printYear . nextMonth . "01"
	return FormatTime(DateAdd(firstDayOfNextMonth, -1, "Days"), "dd")
}

getBlockInfo(fileName) {
	blockInfoMap := Map()
	Xl := ComObject("Excel.Application")
	OnDayGroupDetails := Xl.Workbooks.Open(fileName).Worksheets("Sheet1")
	loop {
		blockCodeReceived := OnDayGroupDetails.Cells(A_Index + 3, 1).Text
		blockNameReceived := OnDayGroupDetails.Cells(A_Index + 3, 2).Text
		if (blockCodeReceived = "" || blockCodeReceived = "Group StayOver") {
			break
		}
		blockInfoMap[blockNameReceived] := blockCodeReceived
	}
	Xl.Workbooks.Close()
	Xl.Quit()
	return blockInfoMap
}

groupArr() {
	groups := []
	Loop {
		blocks := InputBox("请输入blockcode，按下取消或Esc退出", "Arr Group Rooming Lists")
		if (blocks.Result = "Cancel" || blocks.Value = "") {
			break
		}
		groups.push(blocks.Value)
	}
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	For g in groups {
		arrivingGroups(g, g)
	}
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
}

groupArrAuto(blockInfo) {
	BlockInput true
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	for blockName, blockCode in blockInfo {
		arrivingGroups(blockCode, blockName)
	}
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	BlockInput false
}

; hotkeys
; ^F11:: ReportMasterMain()
; F12:: cleanReload()	; use 'Reload' for script reset
; ^F12:: ExitApp	; use 'ExitApp' to kill script