#SingleInstance Force
CoordMode "Mouse", "Screen"

testUI := Gui("","test")
dt := testUI.AddDateTime()
testUI.AddButton("w100","run").OnEvent("Click", query)
testUI.AddButton("w100","info check").OnEvent("Click", checkInfo)

testUI.Show()

flightInfoFormat := [
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

query(*){
	schedulePath := "\\10.0.2.13\fd\25-FEDEX\Schedule 2024\202401 Schedule.xls"
	Xl := ComObject("Excel.Application")
	schdBook := Xl.Workbooks.Open(schedulePath)
	
	global flightDetails := getQueryDateFlights(schdBook)
	
	Xl.Quit()
}

getQueryDateFlights(schdBook) {
	row := 4
	dateQuery := (FormatTime(dt.Value, "MM/dd"))
	sheetCount := schdBook.Worksheets.Count
	inboundFlights := []

	; find query sheet
	loop sheetCount {
		curSheet := schdBook.Worksheets(Format("Sheet{1}", A_Index))
		if (curSheet.Cells(3,1).Text = dateQuery) {
			
			; push each inbound as a Map to inboundFlights
			lastRow := curSheet.Cells(curSheet.Rows.Count, "A").End(-4162).Row
			loop lastRow {
				flightInfoMap := Map()
				loop flightInfoFormat.Length {
					flightInfoMap[flightInfoFormat[A_Index]] := curSheet.Cells(row, A_Index).Text
				}
				inboundFlights.Push(flightInfoMap)
			}
		}
	}
	return inboundFlights
}

checkInfo(*){
	msgbox(flightDetails[1]["flightIn1"] . flightDetails[1]["flightIn2"])
}