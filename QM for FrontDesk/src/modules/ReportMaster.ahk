#Include "../../../Lib/QM for FrontDesk/reports.ahk"
#Include "../../../Lib/Classes/utils.ahk"

class ReportMaster {
	static description := "报表保存 - Report Master"
	static popupTitle := "Report Master"
	; static reportIndex := [
	; 	; [report name, report saving func]
	; 	["1  - Guest INH Complimentary", comp],
	; 	["2  - NA02-Manager Flash", mgrFlash],
	; 	["3  - RS05-（前后15天）History & Forecast", hisFor15],
	; 	["4  - RS05-（FO当月）History & Forecast", hisForThisMonth],
	; 	["5  - RS05-（FO次月）History & Forecast", hisForNextMonth],
	; 	["6  - FO01-VIP Arrival (VIP Arr)", vipArr],
	; 	["7  - FO03-VIP DEP", vipDep],
	; 	["8  - FO01-Arrival Detailed", arrAll],
	; 	["9  - FO02-Guests INH by Room", inhAll],
	; 	["10 - FO03-Departures", depAll],
	; 	["11 - FO11-Credit Limit", creditLimit],
	; 	["12 - FO13-Package Forecast（仅早餐）", bbf],
	; 	["13 - Rooms-housekeepingstatus", rooms],
	; 	["14 - HK03-OOO", ooo],
	; 	["15 - Group Rooming List", groupRoom],
	; 	["16 - Group In House", groupInh],
	; 	["17 - FO08-No Show", noShow],
	; 	["18 - Cancellations", cancel],
	; 	["19 - Guest In House w/o Due Out(VIP INH) ", vipInh]
	; ]
	static reportList := {
		onr: [{
			searchStr: "%complimentary",
			name: "Guest INH Complimentary",
			saveFn: comp
		}, {
			searchStr: "FI01",
			name: "NA02-Manager Flash",
			saveFn: mgrFlash
		}, {
			searchStr: "RS05",
			name: "RS05-（前后15天）History & Forecast",
			saveFn: hisFor15
		}, {
			searchStr: "RS05",
			name: "RS05-（FO当月）History & Forecast",
			saveFn: hisForThisMonth
		}, {
			searchStr: "RS05",
			name: "RS05-（FO次月）History & Forecast",
			saveFn: hisForNextMonth
		}, {
			searchStr: "FO01-VIP",
			name: "VIP Arrival (VIP Arr)",
			saveFn: vipArr
		}, {
			searchStr: "FO03",
			name: "FO03-VIP DEP",
			saveFn: vipDep
		}, {
			searchStr: "FO01",
			name: "FO01-Arrival Detailed",
			saveFn: arrAll
		}, {
			searchStr: "FO02",
			name: "FO02-Guests INH by Room",
			saveFn: inhAll
		}, {
			searchStr: "FO03",
			name: "FO03-Departures",
			saveFn: depAll
		}, {
			searchStr: "FO11",
			name: "FO11-Credit Limit",
			saveFn: creditLimit
		}, {
			searchStr: "FO13",
			name: "FO13-Package Forecast（仅早餐）",
			saveFn: bbf
		}, {
			searchStr: "%hkroomstatusperroom",
			name: "Rooms-housekeepingstatus",
			saveFn: rooms
		}, {
			searchStr: "HK03",
			name: "HK03-OOO",
			saveFn: ooo
		}, {
			searchStr: "gprpmlist",
			name: "Group Rooming List",
			saveFn: groupRoom
		}, {
			searchStr: "grpinhouse",
			name: "Group In House",
			saveFn: groupInh
		}, {
			searchStr: "FO08",
			name: "FO08-No Show",
			saveFn: noShow
		}, {
			searchStr: "rescancel",
			name: "Cancellations",
			saveFn: cancel
		}, {
			searchStr: "%guestinhw",
			name: "Guest In House w/o Due Out(VIP INH)",
			saveFn: vipInh
		},
		],
		misc: [{
			searchStr: "Wshgz_special",
			name: "Specials - 当天水果5 报表",
			saveFn: special
		}, {
			searchStr: "GRPRMLIST",
			name: "Group Arrival - 当天预抵团单",
			saveFn: this.groupArrivingOptions
		},
		]
	}


	static USE() {
		RM := Gui("+Resize", this.popupTitle)

		tab3 := RM.AddTab3("w300 h600", ["夜班报表", "其他报表"])
		tab3.UseTab(1)
		RM.AddText("xp+10 yp+35", "请选择需要保存的报表，日期设定为夜审后操作。`n`n（报表将保存至 开始菜单-文档）。")
		overNightReports := []
		; loop this.reportList.onr.Length {
		; 	overNightReports.Push(
		; 		RM.AddCheckbox("Checked h15 xp y+12", this.reportList.onr[A_Index].name)
		; 	)
		; }
		for report in this.reportList.onr {
			overNightReports.Push(
				RM.AddCheckbox("Checked h15 xp y+12", report.name).OnEvent("Click", checkIsSelectedAll)
			)
		}

		tab3.UseTab(2)
		miscReports := []
		; loop this.reportList.misc.Length {
		; 	miscReports.Push(
		; 		RM.AddRadio("h15 xp y+12", this.reportList.misc[A_Index].name)
		; 	)
		; }
		for report in this.reportList.misc {
			miscReports.Push(
				RM.AddRadio("h15 xp y+12", report.name)
			)
		}
		tab3.UseTab()

		footer := [
			RM.AddCheckbox("vselectAllBtn Checked h30 x15", "全选")
			RM.AddDropDownList("vfileType h30 w80 x+140 Choose1", ["PDF", "XML", "TXT", "XLS"])
			RM.AddButton("h30 w80 x+10", "保存报表").OnEvent("Click", saveReports),
		]

		RM.Show()

		selectAllBtn := Interface.getCtrlByName("selectAllBtn", footer)
		fileType := Interface.getCtrlByName("fileType", footer)

		checkIsSelectedAll(*) {
			selectAllBtn.Value := JSA.every((item) => item.Value = 1, overNightReports) ? 1 : 0
		}

		selectAll(*) {
			loop overNightReports.Length {
				overNightReports[A_Index].Value := overNightReports[A_Index].Value = 1 ? 0 : 1
			}
		}

		saveReports(*) {
			savedReport := ""
			if (tab3.Value = 1) {
				if (JSA.every((item) => item.Value = 0), overNightReports) {
					return
				}
				; loop overNightReports.Length {
				; 	if (JSA.every((item) => item.Value = 0), overNightReports) {
				; 		return
				; 	}
				; 	if (overNightReports[A_Index].Value = 1) {
				; 		if (this.reportList.onr[A_Index].name = "RS05-（FO次月）History & Forecast") {
				; 			if ((this.getLastDay() - A_DD) > 5) {
				; 				continue
				; 			}
				; 		}
				; 		reportFiling(this.reportList.onr[A_Index].saveFn, fileType)
				; 		savedReport .= this.reportList.onr[A_Index].name . "`n"
				; 	}
				; }
				for checkbox in overNightReports {
					if (checkbox.Value = 1) {
						if (this.reportList.onr[A_Index].name = "RS05-（FO次月）History & Forecast") {
							if ((this.getLastDay() - A_DD) > 5) {
								continue
							}
						}
						reportFiling(this.reportList.onr[A_Index].saveFn, fileType)
						savedReport .= this.reportList.onr[A_Index].name . "`n"
					}
				}
				if (JSA.every((item) => item.Value = 1, overNightReports)) {
					savedReport := "夜班报表"
				}
			} else {
				; loop miscReports.Length {
				; 	if (miscReports[A_Index].Value = 1) {
				; 		reportFiling(this.reportList.misc[A_Index].saveFn, fileType)
				; 		savedReport := this.reportList.misc[A_Index].name
				; 	}
				; }
				for checkbox in miscReports {
					if (checkbox.Value = 1) {
						reportFiling(this.reportList.misc[A_Index].saveFn, fileType)
						savedReport := this.reportList.misc[A_Index].name
					}
				}
			}
			Sleep 1000
			this.openMyDocs(savedReport)
		}
	}

	; 	reportMsg := Format("
	; 	(
	; 	请输入对应的报表编号（默认为夜审后操作）。
	; 	（报表将保存至 开始菜单-文档）。

	; 	夜班报表：
	; 	{1}
	; 	666 - 保存以上所有报表（执行时间约8分钟，期间请勿操作电脑）

	; 	其他：
	; 	garr - 保存当天团单
	; 	sp - 保存当天水果5（Excel格式）
	; 	)", this.getReportListStr(this.reportIndex))
	; 	reportSelector := InputBox(reportMsg, this.popupTitle, "h600", "666")
	; 	if (reportSelector.Result = "Cancel") {
	; 		Utils.cleanReload(winGroup)
	; 	}

	; 	if (isNumber(reportSelector.Value)) {
	; 		if (reportSelector.Value > 0 && reportSelector.Value <= this.reportIndex.Length) {
	; 			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	; 			reportName := this.reportIndex[reportSelector.Value][1]
	; 			this.reportIndex[reportSelector.Value][2]()
	; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	; 			Sleep 1000
	; 			this.openMyDocs(reportName)
	; 		} else if (reportSelector.Value = 666) {
	; 			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	; 			loop this.reportIndex.Length {
	; 				if (A_Index = 5) {
	; 					if ((this.getLastDay() - A_DD) > 5) {
	; 						continue
	; 					}
	; 				}
	; 				this.reportIndex[A_Index][2]()
	; 				Sleep 3000
	; 			}
	; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	; 			Sleep 1000
	; 			this.openMyDocs("夜班报表")
	; 		} else {
	; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	; 			MsgBox("请选择表中的指令", this.popupTitle)
	; 		}
	; 	} else {
	; 		if (StrLower(reportSelector.Value) = "garr") {
	; 			reportName := "当天 Arrival 团单"
	; 			onDayGroup := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{2}\{2}{3}\{1}Group ARR&DEP.xlsx", today, A_Year, A_MM)
	; 			blockInfo := this.getBlockInfo(onDayGroup)
	; 			blockInfoText := ""
	; 			for blockName, blockCode in blockInfo {
	; 				blockInfoText .= Format("{1}：{2}`n", blockName, blockCode)
	; 			}
	; 			rmListSaver := MsgBox(Format("
	; 				(
	; 				请确认当天Arrival 团队信息：

	; 				{1}

	; 				是(Y)：自动保存上述团队团单
	; 				否(N)：手动录入block code保存团单
	; 				取消：退出脚本
	; 				)", blockInfoText), this.popupTitle, "YesNoCancel")
	; 			if (rmListSaver = "Yes") {
	; 				this.saveGroupAll(blockinfo)
	; 			} else if (rmListSaver = "No") {
	; 				this.saveGroupRmList()
	; 			} else {
	; 				Utils.cleanReload(winGroup)
	; 			}
	; 			Sleep 1000
	; 			this.openMyDocs(reportName)
	; 		} else if (StrLower(reportSelector.Value) = "sp") {
	; 			reportName := today . "水果5"
	; 			special(today)
	; 			Sleep 1000
	; 			this.openMyDocs(reportName)
	; 		} else {
	; 			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	; 			MsgBox("请选择表中的指令", this.popupTitle)
	; 		}
	; 	}
	; }

	; static getReportListStr(reportList) {
	; 	list := ""
	; 	loop reportList.Length {
	; 		list .= Format("{1}`n", reportList[A_Index][1])
	; 	}
	; 	return list
	; }

	static openMyDocs(reportName) {
		WinSetAlwaysOnTop false
		myText := "已保存报表：`n`n" . reportName . "`n是否打开所在文件夹? "
		openFolder := MsgBox(myText, this.popupTitle, "OKCancel")
		if (openFolder = "OK") {
			Run A_MyDocuments
		} else {
			Utils.cleanReload(winGroup)
		}
	}

	static getLastDay() {
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

	static groupArrivingOptions() {
		reportName := "当天 Arrival 团单"
		onDayGroup := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{2}\{2}{3}\{1}Group ARR&DEP.xlsx", today, A_Year, A_MM)
		blockInfo := this.getBlockInfo(onDayGroup)
		blockInfoText := ""
		for blockName, blockCode in blockInfo {
			blockInfoText .= Format("{1}：{2}`n", blockName, blockCode)
		}
		rmListSaver := MsgBox(Format("
			(	
			请确认当天Arrival 团队信息：

			{1}

			是(Y)：自动保存上述团队团单
			否(N)：手动录入block code保存团单
			取消：退出脚本
			)", blockInfoText), this.popupTitle, "YesNoCancel")
		if (rmListSaver = "Yes") {
			this.saveGroupAll(blockinfo)
		} else if (rmListSaver = "No") {
			this.saveGroupRmList()
		} else {
			Utils.cleanReload(winGroup)
		}
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

	static saveGroupRmList() {
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
		BlockInput true
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		For g in groups {
			if (InStr(g, "=")) {
				g := StrSplit(g, "=")
				arrivingGroups(g[1], g[2])
			} else {
				arrivingGroups(g, g)
			}
		}
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
		BlockInput false
	}

	static saveGroupAll(blockInfo) {
		BlockInput true
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		for blockName, blockCode in blockInfo {
			arrivingGroups(blockCode, blockName)
		}
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
		BlockInput false
	}
}