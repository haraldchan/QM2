; reminder: Y-pos needs to minus 20
; Main runs the main process
; #Include "%A_ScriptDir%\Lib\reports.ahk"
; #Include "%A_ScriptDir%\Lib\utils.ahk"
#Include "../Lib/reports.ahk"
#Include "../Lib/utils.ahk"
today := FormatTime(A_Now, "yyyyMMdd")

ReportMasterMain(params*) {
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"

	prompt := "
	(
	请输入对应的报表编号（默认为夜审后操作）。
	（报表将保存至 开始菜单-文档）。

	1  - Guest INH Complimentary
	2  - NA02-Manager Flash
	3  - RS05-（前后15天）History & Forecast
	4  - RS05-（FO当月）History & Forecast
	5  - RS05-（FO次月）History & Forecast
	6  - FO01-VIP Arrival (VIP Arr)
	7  - Guest In House w/o Due Out 
	8  - FO03-VIP DEP
	9  - FO01-Arrival Detailed
	10 - FO02-Guests INH by Room
	11 - FO03-Departures
	12 - FO11-Credit Limit
	13 - FO13-Package Forecast（仅早餐）
	14 - Rooms-housekeepingstatus
	15 - HK03-OOO
	16 - Group Rooming List
	17 - Group In House
	18 - FO08-No Show
	19 - Cancellations

	garr - 保存当天团单
	sp - 保存当天水果5（Excel格式）
	666 - 保存以上所有报表（执行时间约8分钟，期间请勿操作电脑）`n
	)"

	reportSelector := InputBox(prompt, "ReportMaster", "h600", "666")
	if (reportSelector.Result = "Cancel") {
		cleanReload()
	}
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	; report selector
	switch reportSelector.value, "Off" {
		; { C-shift reports
		Case "1":
			reportName := "1  - Guest INH Complimentary"
			comp()
			Sleep 300
			openMyDocs(reportName)

		Case "2":
			reportName := "2  - NA02-Manager Flash"
			mgrFlash()
			Sleep 300
			openMyDocs(reportName)

		Case "3":
			reportName := "3  - （前后15天）RS05-History & Forecast"
			hisFor15()
			Sleep 300
			openMyDocs(reportName)

		Case "4":
			reportName := "4  - RS05-（FO当月）History & Forecast"
			hisForThisMonth()
			Sleep 300
			openMyDocs(reportName)

		Case "5":
			reportName := "5  - RS05-（FO次月）History & Forecast"
			hisForNextMonth()
			Sleep 300
			openMyDocs(reportName)

		Case "6":
			reportName := "6  - FO01-VIP Arrival (VIP Arr)"
			vipArr()
			Sleep 300
			openMyDocs(reportName)

		Case "7":
			reportName := "7  - Guest In House w/o Due Out"
			vipInh()
			Sleep 300
			openMyDocs(reportName)

		Case "8":
			reportName := "8  - FO03-VIP DEP"
			vipDep()
			Sleep 300
			openMyDocs(reportName)

		Case "9":
			reportName := "9  - FO01-Arrival Detailed"
			arrAll()
			Sleep 300
			openMyDocs(reportName)

		Case "10":
			reportName := "10 - FO02-Guests INH by Room"
			inhAll()
			Sleep 300
			openMyDocs(reportName)

		Case "11":
			reportName := "11 - FO03-Departures"
			depAll()
			Sleep 300
			openMyDocs(reportName)

		Case "12":
			reportName := "12 - FO11-Credit Limit"
			creditLimit()
			Sleep 300
			openMyDocs(reportName)

		Case "13":
			reportName := "13 - FO13-Package Forecast（仅早餐）"
			bbf()
			Sleep 300
			openMyDocs(reportName)

		Case "14":
			reportName := "14 - Rooms-housekeepingstatus"
			rooms()
			Sleep 300
			openMyDocs(reportName)

		Case "15":
			reportName := "15 - HK03-OOO"
			ooo()
			Sleep 300
			openMyDocs(reportName)

		Case "16":
			reportName := "16 - Group Rooming List"
			groupRoom()
			Sleep 300
			openMyDocs(reportName)

		Case "17":
			reportName := "17 - Group In House"
			groupInh()
			Sleep 300
			openMyDocs(reportName)

		Case "18":
			reportName := "18 - FO08-No Show"
			noShow()
			Sleep 300
			openMyDocs(reportName)

		Case "19":
			reportName := "19 - Cancellations"
			cancel()
			Sleep 300
			openMyDocs(reportName)

		Case "666":
			; save all
			comp()
			mgrFlash()
			Sleep 5000
			hisFor15()
			Sleep 5000
			hisForThisMonth()
			Sleep 8000
			; save next month when today is the last 5 day in current month
			preAuditDate := DateAdd(A_Now, -1, "Days")
			preAuditMonth := FormatTime(preAuditDate, "MM")
			preAuditYear := FormatTime(preAuditDate, "yyyy")
			nextMonth := preAuditMonth = 12 ? 1 : preAuditMonth + 1
			if (nextMonth < 10) {
				nextMonth := Format("0{1}", nextMonth)
			}
			printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
			firstDayOfNextMonth := Format("{2}{1}01", nextMonth, printYear)
			dateLast := FormatTime(DateAdd(firstDayOfNextMonth, -1, "Days"), "dd")
			if (dateLast - A_DD < 5) {
				hisForNextMonth()
				Sleep 5000
			}
			vipArr()
			Sleep 5000
			vipInh()
			Sleep 25000
			vipDep()
			Sleep 5000
			arrAll()
			Sleep 5000
			inhAll()
			Sleep 5000
			depAll()
			Sleep 5000
			creditLimit()
			Sleep 5000
			bbf()
			Sleep 5000
			rooms()
			Sleep 5000
			ooo()
			Sleep 5000
			groupRoom()
			Sleep 5000
			groupInh()
			Sleep 5000
			noShow()
			Sleep 5000
			cancel()
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			Sleep 5000
			openMyDocs("全部夜班报表")
			; }

		Case "sp":
			reportName := Format("{1} 水果5", today)
			special(today)
			Sleep 300
			openMyDocs(reportName)

		Case "garr":
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
			reportName := "团单"
			fileName := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{1}Group ARR&DEP.xlsx")
			blockInfo := getBlockInfo()
			blockInfoText := ""
			for blockName, blockCode in blockInfo {
				blockInfoText .= Format("{1}:`t`t{2}`n", blockName, blockCode)
			}
			RmListSaver := MsgBox(Format("
			(	
			请确认当天Arrival 团队信息：

			{1}

			是(Y)：自动保存上述团队团单
			否(N)：手动录入block code保存团单
			取消：退出脚本
			)", blockInfoText), "ReportMaster", "YesNoCancel")
			if (RmListSaver = "Yes") {
				groupArrAuto(blockinfo)
			} else if (RmListSaver = "No") {
				groupArr()
			} else {
				cleanReload()
			}
			Sleep 300
			openMyDocs(reportName)

		Default:
			MsgBox("请选择表中的指令", "ReportMaster")
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	}

}

openMyDocs(reportName) {
	WinSetAlwaysOnTop false
	myText := Format("已保存报表：{1}·`n`n是否打开所在文件夹? ", reportName)
	openFolder := MsgBox(myText, "ReportMaster", "OKCancel")
	if (openFolder = "OK") {
		Run A_MyDocuments
	} else {
		cleanReload()
	}
}

getBlockInfo() {
	blockInfoObj := {}
	Xl := ComObject("Excel.Application")
	fileName := Format("\\10.0.2.13\fd\9-ON DAY GROUP DETAILS\{1}Group ARR&DEP.xlsx", today)
	info := Xl.Workbooks.Open(fileName).Worksheets("Sheet1")
	row := 4
	loop {
		blockCodeReceived := info.Cells(row, 1)
		blockNameReceived := info.Cells(row, 2)
		blockInfoObj.blockNameReceived := blockCodeReceived
		if (blockCodeReceived = "" || blockCodeReceived = "GROUP STAYOVER") {
			break
		}
		row++
	}
	return blockInfoObj
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
	openMyDocs("当日Arrival团单")
}

; hotkeys
;^F11:: ReportMasterMain()
;F12:: cleanReload()	; use 'Reload' for script reset
;^F12:: ExitApp	; use 'ExitApp' to kill script
