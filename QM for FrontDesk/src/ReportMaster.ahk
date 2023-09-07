; #Include "%A_ScriptDir%\Lib\reports.ahk"
; #Include "%A_ScriptDir%\Lib\utils.ahk"
#Include "../Lib/reports.ahk"
#Include "../Lib/utils.ahk"
; scoped vals
; ReportMaster := {
; 	popupTitle: "Report Master",
; }

; overnightReports := [
; 	[comp, "1  - Guest INH Complimentary"],
; 	[mgrFlash, "2  - NA02-Manager Flash"],
; 	[hisFor15, "3  - RS05-（前后15天）History & Forecast"],
; 	[hisForThisMonth, "4  - RS05-（FO当月）History & Forecast"],
; 	[hisForNextMonth, "5  - RS05-（FO次月）History & Forecast"],
; 	[vipArr, "6  - FO01-VIP Arrival (VIP Arr)"],
; 	[vipDep, "7  - FO03-VIP DEP"],
; 	[arrAll, "8  - FO01-Arrival Detailed"],
; 	[inhAll, "9  - FO02-Guests INH by Room"],
; 	[depAll, "10 - FO03-Departures"],
; 	[creditLimit, "11 - FO11-Credit Limit"],
; 	[bbf, "12 - FO13-Package Forecast（仅早餐）"],
; 	[rooms, "13 - Rooms-housekeepingstatus"],
; 	[ooo, "14 - HK03-OOO"],
; 	[groupRoom, "15 - Group Rooming List"],
; 	[groupInh, "16 - Group In House"],
; 	[noShow, "17 - FO08-No Show"],
; 	[cancel, "18 - Cancellations"],
; 	[vipInh, "19 - Guest In House w/o Due Out(VIP INH) "]
; ]

; ReportMasterMain() {
; 	WinMaximize "ahk_class SunAwtFrame"
; 	WinActivate "ahk_class SunAwtFrame"
; 	reportMsg := "
; 	(
; 	请输入对应的报表编号（默认为夜审后操作）。
; 	（报表将保存至 开始菜单-文档）。

; 	夜班报表：
; 	1  - Guest INH Complimentary
; 	2  - NA02-Manager Flash
; 	3  - RS05-（前后15天）History & Forecast
; 	4  - RS05-（FO当月）History & Forecast
; 	5  - RS05-（FO次月）History & Forecast
; 	6  - FO01-VIP Arrival (VIP Arr)
; 	7  - FO03-VIP DEP
; 	8  - FO01-Arrival Detailed
; 	9  - FO02-Guests INH by Room
; 	10 - FO03-Departures
; 	11 - FO11-Credit Limit
; 	12 - FO13-Package Forecast（仅早餐）
; 	13 - Rooms-housekeepingstatus
; 	14 - HK03-OOO
; 	15 - Group Rooming List
; 	16 - Group In House
; 	17 - FO08-No Show
; 	18 - Cancellations
; 	19 - Guest In House w/o Due Out 
; 	666 - 保存以上所有报表（执行时间约8分钟，期间请勿操作电脑）

; 	其他：
; 	garr - 保存当天团单
; 	sp - 保存当天水果5（Excel格式）
; 	)"
; 	reportSelector := InputBox(reportMsg, ReportMaster.popupTitle, "h600", "666")
; 	if (reportSelector.Result = "Cancel") {
; 		cleanReload()
; 	}
; 	try { ; input value is number
; 		if (reportSelector.Value > 0 && reportSelector.Value <= overnightReports.Length) {
; 			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
; 			reportName := overnightReports[reportSelector.Value][2]
; 			overnightReports[reportSelector.Value][1]()
; 			Sleep 1500
; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; 			openMyDocs(reportName)
; 		} else if (reportSelector.Value = 666) {
; 			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
; 			loop overnightReports.Length {
; 				if (A_Index = 5) {
; 					if ((dateLast() - A_DD) > 5) {
; 						continue
; 					}
; 				}
; 				overnightReports[A_Index][1]()
; 				Sleep 2500
; 			}
; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; 			openMyDocs("夜班报表")
; 		} else {
; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; 			MsgBox("请选择表中的指令", ReportMaster.popupTitle)
; 		}
; 	}
; 	catch { ; input value is string
; 		if (StrLower(reportSelector.Value) = "garr") {
; 			reportName := "当天 Arrival 团单"
; 			onDayGroup := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{2}\{2}{3}\{1}Group ARR&DEP.xlsx", today, A_Year, A_MM)
; 			blockInfo := getBlockInfo(onDayGroup)
; 			blockInfoText := ""
; 			for blockName, blockCode in blockInfo {
; 				blockInfoText .= Format("{1}：{2}`n", blockName, blockCode)
; 			}
; 			RmListSaver := MsgBox(Format("
; 					(	
; 					请确认当天Arrival 团队信息：

; 					{1}

; 					是(Y)：自动保存上述团队团单
; 					否(N)：手动录入block code保存团单
; 					取消：退出脚本
; 					)", blockInfoText), ReportMaster.popupTitle, "YesNoCancel")
; 	if (RmListSaver = "Yes") {
; 		groupArrAuto(blockinfo)
; 	} else if (RmListSaver = "No") {
; 		groupArr()
; 	} else {
; 		cleanReload()
; 	}
; 	Sleep 300
; 	openMyDocs(reportName)
; 		} else if (StrLower(reportSelector.Value) = "sp") {
; 			reportName := today . "水果5"
; 			special(today)
; 			Sleep 300
; 			openMyDocs(reportName)
; 		} else {
; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; 			MsgBox("请选择表中的指令", ReportMaster.popupTitle)
; 		}
; 	}
; }

; openMyDocs(reportName) {
; 	WinSetAlwaysOnTop false
; 	myText := "已保存报表：" . reportName . "`n`n是否打开所在文件夹? "
; 	openFolder := MsgBox(myText, ReportMaster.popupTitle, "OKCancel")
; 	if (openFolder = "OK") {
; 		Run A_MyDocuments
; 	} else {
; 		cleanReload()
; 	}
; }

; dateLast() {
; 	preAuditDate := DateAdd(A_Now, -1, "Days")
; 	preAuditMonth := FormatTime(preAuditDate, "MM")
; 	preAuditYear := FormatTime(preAuditDate, "yyyy")
; 	nextMonth := (preAuditMonth = 12) ? 1 : preAuditMonth + 1
; 	if (nextMonth < 10) {
; 		nextMonth := "0" . nextMonth
; 	}
; 	printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
; 	firstDayOfNextMonth := printYear . nextMonth . "01"
; 	return FormatTime(DateAdd(firstDayOfNextMonth, -1, "Days"), "dd")
; }

; getBlockInfo(fileName) {
; 	blockInfoMap := Map()
; 	Xl := ComObject("Excel.Application")
; 	OnDayGroupDetails := Xl.Workbooks.Open(fileName).Worksheets("Sheet1")
; 	loop {
; 		blockCodeReceived := OnDayGroupDetails.Cells(A_Index + 3, 1).Text
; 		blockNameReceived := OnDayGroupDetails.Cells(A_Index + 3, 2).Text
; 		if (blockCodeReceived = "" || blockCodeReceived = "Group StayOver") {
; 			break
; 		}
; 		blockInfoMap[blockNameReceived] := blockCodeReceived
; 	}
; 	Xl.Workbooks.Close()
; 	Xl.Quit()
; 	return blockInfoMap
; }

; groupArr() {
; 	groups := []
; 	Loop {
; 		blocks := InputBox("
; 		(
; 		请输入blockcode，按下取消或Esc退出
		
; 		如需另外命名pdf 文件，请以“blockcode=文件名”形式输入
; 		如：20230901huaha=华海
; 		)", "Arr Group Rooming Lists")
; 	if (blocks.Result = "Cancel" || blocks.Value = "") {
; 		break
; 	}
; 	groups.push(blocks.Value)
; 	}
; 	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
; 	For g in groups {
; 		if (InStr(g, "=")) {
; 			StrSplit(g, "=")
; 			arrivingGroups(g[1], g[2])
; 		} else {
; 			arrivingGroups(g, g)
; 		}
; 	}
; 	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; }

; groupArrAuto(blockInfo) {
; 	BlockInput true
; 	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
; 	for blockName, blockCode in blockInfo {
; 		arrivingGroups(blockCode, blockName)
; 	}
; 	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
; 	BlockInput false
; }

class ReportMaster {
	static popupTitle := "Report Master"
	static overnightReports := [
		[comp, "1  - Guest INH Complimentary"],
		[mgrFlash, "2  - NA02-Manager Flash"],
		[hisFor15, "3  - RS05-（前后15天）History & Forecast"],
		[hisForThisMonth, "4  - RS05-（FO当月）History & Forecast"],
		[hisForNextMonth, "5  - RS05-（FO次月）History & Forecast"],
		[vipArr, "6  - FO01-VIP Arrival (VIP Arr)"],
		[vipDep, "7  - FO03-VIP DEP"],
		[arrAll, "8  - FO01-Arrival Detailed"],
		[inhAll, "9  - FO02-Guests INH by Room"],
		[depAll, "10 - FO03-Departures"],
		[creditLimit, "11 - FO11-Credit Limit"],
		[bbf, "12 - FO13-Package Forecast（仅早餐）"],
		[rooms, "13 - Rooms-housekeepingstatus"],
		[ooo, "14 - HK03-OOO"],
		[groupRoom, "15 - Group Rooming List"],
		[groupInh, "16 - Group In House"],
		[noShow, "17 - FO08-No Show"],
		[cancel, "18 - Cancellations"],
		[vipInh, "19 - Guest In House w/o Due Out(VIP INH) "]
	]

	static Main() {
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
		reportSelector := InputBox(reportMsg, this.popupTitle, "h600", "666")
		if (reportSelector.Result = "Cancel") {
			cleanReload()
		}
		try { ; input value is number
			if (reportSelector.Value > 0 && reportSelector.Value <= this.overnightReports.Length) {
				WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
				reportName := this.overnightReports[reportSelector.Value][2]
				this.overnightReports[reportSelector.Value][1]()
				Sleep 1500
				WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
				this.openMyDocs(reportName)
			} else if (reportSelector.Value = 666) {
				WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
				loop this.overnightReports.Length {
					if (A_Index = 5) {
						if ((this.dateLast() - A_DD) > 5) {
							continue
						}
					}
					this.overnightReports[A_Index][1]()
					Sleep 2500
				}
				WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
				this.openMyDocs("夜班报表")
			} else {
				WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
				MsgBox("请选择表中的指令", ReportMaster.popupTitle)
			}
		}
		catch { ; input value is string
			if (StrLower(reportSelector.Value) = "garr") {
				reportName := "当天 Arrival 团单"
				onDayGroup := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{2}\{2}{3}\{1}Group ARR&DEP.xlsx", today, A_Year, A_MM)
				blockInfo := this.getBlockInfo(onDayGroup)
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
						)", blockInfoText), this.popupTitle, "YesNoCancel")
			if (RmListSaver = "Yes") {
				this.groupArrAuto(blockinfo)
			} else if (RmListSaver = "No") {
				this.groupArr()
			} else {
				cleanReload()
			}
			Sleep 300
			this.openMyDocs(reportName)
			} else if (StrLower(reportSelector.Value) = "sp") {
				reportName := today . "水果5"
				special(today)
				Sleep 300
				this.openMyDocs(reportName)
			} else {
				WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
				MsgBox("请选择表中的指令", this.popupTitle)
			}
		}
	}

	static openMyDocs(reportName) {
		WinSetAlwaysOnTop false
		myText := "已保存报表：" . reportName . "`n`n是否打开所在文件夹? "
		openFolder := MsgBox(myText, this.popupTitle, "OKCancel")
		if (openFolder = "OK") {
			Run A_MyDocuments
		} else {
			cleanReload()
		}
	}

	static dateLast() {
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
	
	static getBlockInfo(fileName) {
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

	static groupArr() {
		groups := []
		Loop {
			blocks := InputBox("
			(
			请输入blockcode，按下取消或Esc退出
			
			如需另外命名pdf 文件，请以“blockcode=文件名”形式输入
			如：20230901huaha=华海
			)", "Arr Group Rooming Lists")
		if (blocks.Result = "Cancel" || blocks.Value = "") {
			break
		}
		groups.push(blocks.Value)
		}
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		For g in groups {
			if (InStr(g, "=")) {
				StrSplit(g, "=")
				arrivingGroups(g[1], g[2])
			} else {
				arrivingGroups(g, g)
			}
		}
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}

	static groupArrAuto(blockInfo) {
		BlockInput true
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		for blockName, blockCode in blockInfo {
			arrivingGroups(blockCode, blockName)
		}
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
		BlockInput false
	}
}