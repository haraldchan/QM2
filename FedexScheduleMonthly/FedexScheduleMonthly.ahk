config := Format("{1}\config.ini", A_ScriptDir)
path := IniRead(config, "FSM", "schedulePath")
schdRange := IniRead(config, "FSM", "scheduleRange")
tempPath := "\\10.0.2.13\fd\25-FEDEX\Schedule生成处理工具\FedEx Sign In Sheet temp(空模板).xlsx"
saveDir := Format("\\10.0.2.13\fd\25-FEDEX\Sign-In sheet生成({1})", schdRange)
DirCreate saveDir

Xl := ComObject("Excel.Application")
schdData := Xl.Workbooks.Open(path)
schdSheetAmount := schdData.Sheets.Count

sheet := 1
schdRow := 4
tempRow := 3
flightFormat := [
	"tripNum",
	"roomQty",
	"flightIn1",
	"flightIn2",
	"ibDate",
	"ETA",
	"stayHours",
	"obDate",
	"ETD",
	"flightOut1",
	"flightOut2"
]

loop schdSheetAmount {
	template := Xl.Workbooks.Open(tempPath)
	sheetIndex := Format("Sheet{1}", sheet)
	schdData.Worksheets(sheetIndex).Activate
	ciDate := schdData.ActiveSheet.Cells(3, 1).Text
	lastRow := schdData.ActiveSheet.Cells(schdData.ActiveSheet.Rows.Count,"A").End(-4162).Row
	if (ciDate = "") {
		schdData.Close
		template.Close
		Xl.Quit
		break
	}
	if (StrSplit(ciDate, "/")[1] < A_MM) {
		fileDate := Format("{1}{2}", A_Year + 1, StrReplace(ciDate, "/", ""))
	} else {
		fileDate := Format("{1}{2}", A_Year, StrReplace(ciDate, "/", ""))
	}
	TrayTip Format("正在保存：{1} FedEx Sign In Sheet.xlsx", fileDate)

	loop (lastRow - 3) {
		flightInfo := []
		loop flightFormat.Length {
			flightInfo.Push(schdData.Worksheets(sheetIndex).Cells(schdRow, A_Index).Text)
		}
		; flightInfoText := ""
		; 	for f in flightInfo {
		; 		flightInfoText .= Format("{1}：{2}`n", flightFormat[A_Index], flightInfo[A_Index])
		; 	}
		; 	MsgBox(flightInfoText)
		loop flightFormat.Length {
			template.ActiveSheet.Cells(tempRow, A_Index + 4).Value := flightInfo[A_Index]

		}
		schdRow++
		tempRow++
	}
	sheet++
	template.SaveAs(Format("{1}\{2}FedEx Sign In Sheet.xlsx", saveDir, fileDate))
	schdRow := 4
	tempRow := 3
}
Xl.Quit()
checkSavedSchdules := MsgBox("已生成Sign-in Sheet文件。`n是否打开保存所在文件夹查看？", "FedexScheduleMonthly", "OKCancel")
if (checkSavedSchdules = "OK") {
	Run saveDir
}