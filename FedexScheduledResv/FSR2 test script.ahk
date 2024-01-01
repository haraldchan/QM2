#Include "./utils.ahk"

class FedexScheduledReservations {
	static flightInfoItems := [
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
	static resvOnDayTime := IniRead(A_ScriptDir . "\config.ini", "FSR", "resvOnDayTime")
	static bringForwardTime := getBringForwardTime(this.resvOnDayTime)

	static getQueryDateFlights(schedule, date) {
		row := 4
		dateQuery := (FormatTime(date, "MM/dd"))
		sheetCount := schedule.Worksheets.Count
		inboundFlights := []

		; find query sheet
		loop sheetCount {
			curSheet := schedule.Worksheets(Format("Sheet{1}", A_Index))
			if (curSheet.Cells(3, 1).Text = dateQuery) {

				; push each inbound as a Map to inboundFlights(array)
				lastRow := curSheet.Cells(curSheet.Rows.Count, "A").End(-4162).Row
				loop (lastRow - 3) {
					flightInfoMap := Map()
					for item in this.flightInfoItems {
						flightInfoMap[item] := curSheet.Cells(row, A_Index).Text
					}
					inboundFlights.Push(flightInfoMap)
					row++
				}
			}
		}

		return inboundFlights
	}

	static writeReservations(inboundflights) {

		for line in inboundflights {
			loop line["roomQty"] {
				FsrEntry.USE(line, this.bringForwardTime)
			}
		}

		MsgBox("已完成FedEx 预订录入，请抽检以确保准确！", "FedexScheduledReservations")
	}
}

class FsrEntry {
	static USE(inboundFlightLine, bringForwardTime) {
		; { date reformatting
		actualYear := (StrSplit(inboundFlightLine["ibDate"], "/")[1] < A_MM)
			? A_Year + 1
			: A_Year
		schdCiDate := Format("{1}{2}{3}", actualYear, StrSplit(inboundFlightLine["ibDate"], "/")[1], StrSplit(inboundFlightLine["ibDate"], "/")[2])
		schdCoDate := Format("{1}{2}{3}", actualYear, StrSplit(inboundFlightLine["obDate"], "/")[1], StrSplit(inboundFlightLine["obDate"], "/")[2])
		daysActual := getDaysActual(inboundFlightLine["stayHours"])

		pmsCiDate := (StrSplit(inboundFlightLine["ETA"], ":")[1]) < bringForwardTime
			? DateAdd(schdCiDate, -1, "days")
			: schdCiDate
		pmsCoDate := schdCoDate
		comment := Format("RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}day(s), ActualStay: {3}-{4}",
			inboundFlightLine["stayHours"],
			daysActual,
			schdCiDate,
			schdCoDate
		)
		; flightIn := inboundFlightLine["flightIn1"] . inboundFlightLine["flightIn2"]
		; flightOut := inboundFlightLine["flightOut1"] . inboundFlightLine["flightOut2"]
		pmsNts := DateDiff(pmsCoDate, pmsCiDate, "days")
		; reformat to match pms date format
		schdCiDate := FormatTime(schdCiDate, "MMddyyyy")
		schdCoDate := FormatTime(schdCoDate, "MMddyyyy")
		pmsCiDate := FormatTime(pmsCiDate, "MMddyyyy")
		pmsCoDate := FormatTime(pmsCoDate, "MMddyyyy")
		; }
		this.openBooking()
		Sleep 500

		this.profileEntry(inboundFlightLine["flightIn"], inboundFlightLine["tripNum"])
		Sleep 500

		this.dateTimeEntry(pmsCiDate, pmsCoDate, inboundFlightLine["ETA"], inboundFlightLine["ETD"])
		Sleep 500

		this.commentIbdTripEntry(comment, inboundFlightLine["flightIn"], inboundFlightLine["tripNum"])
		Sleep 500

		this.moreFieldsEntry(schdCiDate, schdCoDate, inboundFlightLine["ETA"], inboundFlightLine["ETD"], inboundFlightLine["flightIn"], inboundFlightLine["flightOut"])
		Sleep 500

		if (inboundFlightLine["daysActual"] < pmsNts) {
			this.dailyDetailsEntry(inboundFlightLine["daysActual"])
		}
		Sleep 500

		this.saveBooking()
		Sleep 500
	}

	static openBooking() {
		Sleep 1000
		MouseMove 399, 244
		Sleep 150
		Send "!e"
		Sleep 1000
	}

	static saveBooking() {
		Send "!o"
		Sleep 1000
		Send "!o"
		Sleep 1000
		Send "{Down}"
		Sleep 2000
	}

	static profileEntry(flightIn, tripNumber, initX := 467, initY := 201) {
		MouseMove 467, 201
		Sleep 100
		Click
		Sleep 1000
		MouseMove 442, 264
		Sleep 100
		Click "Down"
		MouseMove 214, 269
		Sleep 100
		Click "Up"
		Sleep 100
		Send "{Backspace}"
		Sleep 100
		Send Format("{Text}{1}  {3}", flightIn, tripNumber)
		Sleep 100
		MouseMove 594, 394
		Sleep 150
		Send "!c"
		Sleep 100
		MouseMove 576, 521
		Sleep 100
		Click
		Sleep 100
		MouseMove 798, 480
		Sleep 100
		Click
		Sleep 100
	}

	static dateTimeEntry(checkin, checkout, ETA, ETD, initX := 323, initY := 506) {
		pmsCiDate := (StrSplit(ETA, ":")[1]) < 10
			? FormatTime(DateAdd(checkin, -1, "days"), "MMddyyyy")
			: FormatTime(checkin, "MMddyyyy")
		pmsCoDate := FormatTime(checkout, "MMddyyyy")
		MouseMove initX + 9, initY - 150 ; 332, 356
		Sleep 100
		; fill-in checkin/checkout
		Click "Down"
		MouseMove initX - 145, initY - 146 ; 178, 360
		Sleep 100
		Click "Up"
		MouseMove initX - 151, initY - 146 ; 172, 360
		Sleep 100
		Send Format("{Text}{1}", pmsCiDate)
		Sleep 100
		MouseMove initX + 2, initY - 108 ; 325, 398
		Sleep 100
		Click
		Sleep 100
		MouseMove initX + 338, initY + 37 ; 661, 543
		Sleep 100
		Click
		MouseMove initX + 313, initY + 37 ; 636, 543
		Sleep 100
		Click
		MouseMove initX + 312, initY + 37 ; 635, 543
		Sleep 100
		Click
		Sleep 100
		Click
		Sleep 100
		MouseMove initX + 12, initY - 101 ; 335, 405
		Sleep 100
		Click "Down"
		MouseMove initX - 141, initY - 97 ; 182, 409
		Sleep 100
		Click "Up"
		MouseMove initX - 116, initY - 91 ; 207, 415
		Sleep 100
		Send Format("{Text}{1}", pmsCoDate)
		Sleep 100
		Send "{Enter}"
		Sleep 100
		loop 3 {
			Send "{Esc}"
			Sleep 200
		}
		; fill in ETA & ETD
		MouseMove initX - 29, initY + 91 ; 294, 597
		Sleep 100
		MouseMove initX - 3, initY + 91 ; 320, 597
		Sleep 100
		Click "Down"
		MouseMove initX - 123, initY + 91 ; 200, 597
		Sleep 100
		Click "Up"
		Sleep 100
		Send Format("{Text}{1}", ETA)
		Sleep 100
		MouseMove initX + 176, initY + 91 ; 499, 597
		Sleep 100
		Click "Down"
		MouseMove initX + 7, initY + 88 ; 330, 594
		Sleep 100
		Click "Up"
		Sleep 100
		Send Format("{Text}{1}", ETD)
		Sleep 100
	}

	static commentIbdTripEntry(comment, flightIn, tripNumber, initX := 622, initY := 589) {
		; select all and re-enter comment
		MouseMove initX, initY ; 622, 596
		Sleep 100
		Click "Down"
		MouseMove initX + 518, initY + 36 ; 1140, 605
		Sleep 100
		Click "Up"
		Sleep 100
		Send "{Backspace}"
		Sleep 100
		Send comment
		Sleep 100
		; fill-in new flight and trip
		MouseMove initX + 307, initY - 35 ; 929, 554
		Sleep 100
		Click 3
		Sleep 100
		Send Format("{Text}{1}  {2}", flightIn, tripNumber)
		Sleep 100
	}

	static moreFieldsEntry(sCheckin, sCheckout, ETA, ETD, flightIn, flightOut, initX := 236, initY := 333) {
		MouseMove initX, initY ; 236, 333
		Sleep 100
		Click
		Sleep 100
		MouseMove initX + 450, initY + 126 ; 686, 459
		Sleep 200
		Click "Down"
		MouseMove initX + 242, initY + 126 ; 478, 459
		Sleep 100
		Click "Up"
		Sleep 100
		Send Format("{Text}{1}", flightIn)
		Sleep 100
		MouseMove initX + 436, initY + 170 ; 672, 503
		Sleep 100
		Click "Down"
		MouseMove initX + 287, initY + 170 ; 523, 503
		Sleep 100
		Click "Up"
		Send Format("{Text}{1}", sCheckin)
		Sleep 100
		MouseMove initX + 449, initY + 193 ; 685, 526
		Sleep 100
		Click "Down"
		MouseMove initX + 186, initY + 188 ; 422, 521
		Sleep 100
		Click "Up"
		Send Format("{Text}{1}", ETA)
		Sleep 100
		MouseMove initX + 686, initY + 128 ; 922, 461
		Sleep 100
		Click "Down"
		MouseMove initX + 468, initY + 126 ; 704, 459
		Sleep 100
		Click "Up"
		Send Format("{Text}{1}", flightOut)
		Sleep 100
		MouseMove initX + 681, initY + 170 ; 917, 503
		Sleep 100
		Click "Down"
		MouseMove initX + 401, initY + 174 ; 637, 507
		Sleep 100
		Click "Up"
		Sleep 100
		Send Format("{Text}{1}", sCheckout)
		Sleep 100
		MouseMove initX + 686, initY + 191 ; 922, 524
		Sleep 100
		Click "Down"
		MouseMove initX + 404, initY + 190 ; 640, 523
		Sleep 100
		Click "Up"
		Sleep 100
		Send Format("{Text}{1}", ETD)
		Sleep 100
		MouseMove initX + 605, initY + 347 ; 841, 680
		Sleep 100
		Click
		Sleep 100
	}

	static dailyDetailsEntry(daysActual, initX := 372, initY := 524) {
		MouseMove initX, initY ; 372, 524
		Sleep 100
		Click
		Sleep 100
		Send "!d"
		Sleep 1500
		loop daysActual {
			Send "{Down}"
			Sleep 200
		}
		Send "!e"
		Sleep 100
		loop 4 {
			Send "{Tab}"
			Sleep 200
		}
		Send "{Text}NRR"
		Sleep 100
		Send "!o"
		loop 3 {
			Send "{Enter}"
			Sleep 200
		}
		Sleep 300
		Send "!o"
		Sleep 100
		loop 5 {
			Send "{Escape}"
			Sleep 200
		}
		Sleep 100
	}


}