FsmMain() {
	config := Format("{1}\config.ini", A_ScriptDir)
	path := IniRead(config, "FSM", "schedulePath")
	schdRange := IniRead(config, "FSM", "scheduleRange")
	tempPath := "\\10.0.2.13\fd\25-FEDEX\Schedule生成处理工具\FedEx Sign In Sheet temp(空模板).xlsx"
	saveDir := Format("\\10.0.2.13\fd\25-FEDEX\Sign-In sheet生成({1})", schdRange)
	try {
		DirCreate saveDir
	} catch {
		MsgBox("指定日期范围内不可包含以下字符: \ / “ * < > | ", "FedexScheduleMonthly")
	}

	Xl := ComObject("Excel.Application")
	schdData := Xl.Workbooks.Open(path)
	schdSheetAmount := schdData.Sheets.Count
	sheet := 1
	schdRow := 4
	tempRow := 3

	loop schdSheetAmount {
		template := Xl.Workbooks.Open(tempPath)
		sheetIndex := Format("Sheet{1}", sheet)
		schdData.Worksheets(sheetIndex).Activate
		ciDate := schdData.ActiveSheet.Cells(3, 1).Text
		lastRow := schdData.ActiveSheet.Cells(schdData.ActiveSheet.Rows.Count, "A").End(-4162).Row
		if (ciDate = "") {
			schdData.Close()
			template.Close()
			break
		}
		fileDate := StrSplit(ciDate, "/")[1] < A_MM
			? fileDate := Format("{1}{2}", A_Year + 1, StrReplace(ciDate, "/", ""))
			: fileDate := Format("{1}{2}", A_Year, StrReplace(ciDate, "/", ""))
		loop (lastRow - 3) {
			loop 11 {
				template.ActiveSheet.Cells(tempRow, A_Index + 4).Value := schdData.Worksheets(sheetIndex).Cells(schdRow, A_Index).Text
			}
			schdRow++
			tempRow++
		}
		sheet++
		template.SaveAs(Format("{1}\{2}FedEx Sign In Sheet.xlsx", saveDir, fileDate))
		TrayTip Format("正在保存：{1} FedEx Sign In Sheet.xlsx", fileDate)
		schdRow := 4
		tempRow := 3
	}
	Xl.Quit()
	checkSavedSchdules := MsgBox("已生成Sign-in Sheet文件。`n是否打开保存所在文件夹查看？", "FedexScheduleMonthly", "OKCancel")
	if (checkSavedSchdules = "OK") {
		Run saveDir
	}
}