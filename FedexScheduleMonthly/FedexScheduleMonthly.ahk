path := IniRead("%A_ScriptDir%\config.ini", "FSM", "schedulePath")
range := IniRead("%A_ScriptDir%\config.ini", "FSM", "scheduleRange")
tempPath := "\\10.0.2.13\fd\25-FEDEX\Schedule生成处理工具\FedEx Sign In Sheet temp(空模板).xlsx"

if (DirExist("\\10.0.2.13\fd\25-FEDEX\Sign-In sheet生成\%range%") = false) {
	DirCreate "\\10.0.2.13\fd\25-FEDEX\Sign-In sheet生成\%range%"
}

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
	sheetIndex := Format("Sheet{1}", A_Index)
	template := Xl.Workbooks.Open(tempPath)
	ciDate := schdData.Worksheets(sheetIndex).Cells(3, 1).Text
	lastRow := schdData.Worksheets(sheetIndex).UsedRange.Rows.Count
	if (StrSplit(ciDate, "/")[1] < A_MM) {
		fileDate := Format("{1}{2}", A_Year + 1, StrReplace(ciDate, "/", ""))
	} else {
		fileDate := Format("{1}{2}", A_Year, StrReplace(ciDate, "/", ""))
	}
	TrayTip "正在保存：%fileDate% FedEx Sign In Sheet.xlsx"

	loop (lastRow - 3) {
		flightInfo := []
		loop flightFormat.Length {
			flightInfo[A_Index] := schdData.ActiveSheet.Cells(schdRow, A_Index).Text
		}
		loop flightFormat.Length {
			template.ActiveSheet.Cells(tempRow, A_Index + 4) := flightInfo[A_Index]
		}
		schdRow++
		tempRow++
	}
	saveName := "\\10.0.2.13\25-FEDEX\Sign-In sheet生成(%range%)\%ciDate%FedEx Sign In Sheet.xlsx"
	template.SaveAs(saveName)
	template.close
	schdRow := 4
	tempRow := 3
}

checkSavedSchdules := MsgBox("已生成Sign-in Sheet文件。`n是否打开保存所在文件夹查看？", "FedexScheduleMonthly", "OKCancel")
if (checkSavedSchdules = "OK") {
	Run "\\10.0.2.13\fd\25-FEDEX\Sign-In sheet生成\%range%"
}