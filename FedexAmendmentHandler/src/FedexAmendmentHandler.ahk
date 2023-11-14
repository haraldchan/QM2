#Include "%A_ScriptDir%\src\utils.ahk"

class FedexAmendmentHandler {
	static path := A_ScriptDir . "\src\FedExPasteTemplate.xls"

	static getAmendmentInfo(){
		amendmentMap := Map()

		Xl := ComObject("Excel.Application")
		Xlbook := Xl.Workbooks.Open(this.path)
		amendment := Xlbook.Worksheets("Sheet1") 
		lastRow := amendment.Cells(Xlbook.ActiveSheet.Rows.Count, "A").End(-4162).Row

		resvType := StrSplit(amendment.Cells(9,1).Text, " ")[1]
		readStartRow := (resvType = "ADD") ? 23 : 28

		trackingNum := SubStr(amendment.Cells(7,1).Text, 14)
		tripNum := StrSplit(amendment.Cells(8,1).Text, " ")[2] . "/" . StrSplit(amendment.Cells(8,1).Text, " ")[4]

		roomQty := StrSplit(amendment.Cells(readStartRow,1).Text, " ")[2]

		inBoundInfo := StrSplit(amendment.Cells(readStartRow+1,1).Text, " ")
		flightIn := inBoundInfo[3] . inBoundInfo[4]
		ibDate := getReformatDate(inBoundInfo[7])
		ETA := get24Hr(inBoundInfo[8], inBoundInfo[9])

		outBoundInfo := StrSplit(amendment.Cells(readStartRow+2,1).Text, " ")
		flightOut := outBoundInfo[3] . outBoundInfo[4]
		obDate := getReformatDate(outBoundInfo[7])
		ETD := get24Hr(outBoundInfo[8], outBoundInfo[9])

		stayHours := getStayHours(ibDate, ETA, obDate, ETD)
		daysActual := getDaysActual(stayHours)

		staffInfoText := amendment.Cells(lastRow-3, 1).Text
		staffInfo := getStaffNames(staffInfoText)

		Xl.Quit()

		amendmentMap["resvType"] := resvType
		amendmentMap["trackingNum"] := trackingNum
		amendmentMap["tripNum"] := tripNum
		amendmentMap["roomQty"] := roomQty
		amendmentMap["flightIn"] := flightIn
		amendmentMap["ibDate"] := ibDate
		amendmentMap["ETA"] := ETA
		amendmentMap["flightOut"] := flightOut
		amendmentMap["obDate"] := obDate
		amendmentMap["ETD"] := ETD
		amendmentMap["stayHours"] := stayHours
		amendmentMap["daysActual"] := daysActual
		amendmentMap["staffInfo"] := staffInfo

		; info msgbox
		msgInfo := Format("
			(
			房间数量：     {1}
			入住人员： 
			{2}

			Trip No. :    {3}
			FlightIn :     {4}
			Inbound:    {5} 
			ETA:           {6}

			FlightOut:    {7}
			Outbound:  {8}
			ETD:           {9}

			在住时长:      {10}
			计费天数:      {11}
			)",roomQty, staffInfoText, tripNum, flightIn, ibDate, ETA, flightOut, obDate, ETD, stayHours, daysActual)
		msgbox(msgInfo)

		return amendmentMap
	}

	
}