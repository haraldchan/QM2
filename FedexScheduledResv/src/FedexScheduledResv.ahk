#Include "./FSR_Utils.ahk"

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
	static resvOnDayTime := IniRead(A_ScriptDir . "\src\config.ini", "FSR", "resvOnDayTime")
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
		WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
		WinMaximize "ahk_class SunAwtFrame"
		BlockInput true

		for line in inboundflights {
			loop line["roomQty"] {
				FsrEntry.USE(line, this.bringForwardTime)
			}
		}
		WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
		BlockInput false

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
		comment := Format("{Text}RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}day(s), ActualStay: {3}-{4}",
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
		
		this.profileEntry(inboundFlightLine["flightIn"], inboundFlightLine["tripNum"])

		this.dateTimeEntry(pmsCiDate, pmsCoDate, inboundFlightLine["ETA"], inboundFlightLine["ETD"])

		this.commentIbdTripEntry(comment, inboundFlightLine["flightIn"], inboundFlightLine["tripNum"])

		this.moreFieldsEntry(schdCiDate, schdCoDate, inboundFlightLine["ETA"], inboundFlightLine["ETD"], inboundFlightLine["flightIn"], inboundFlightLine["flightOut"])

		if (daysActual < pmsNts) {
			this.dailyDetailsEntry(daysActual)
		}

		this.saveBooking()
	}

	static openBooking() {
		waitLoading()
		MouseMove 399, 244
		Sleep 150
		Send "!e"
		waitLoading()
	}

	static saveBooking() {
		CoordMode "Pixel", "Screen"
		Send "!o"
		waitLoading()
		loop {
			waitLoading()
			if (PixelGetColor(610, 330) = "0x99B4D1") { 
				break
			}
			if (A_Index = 20) {
				WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
				BlockInput false
				MsgBox("已停止。")
				cleanReload()
			}
		}
		Send "!o"
		waitLoading()
		Send "{Down}"
		waitLoading()
	}

	static profileEntry(flightIn, tripNumber, initX := 467, initY := 201) {
		MouseMove 467, 221
		waitLoading()
		Click
		waitLoading()
		MouseMove 442, 284
		waitLoading()
		Click "Down"
		MouseMove 214, 289
		waitLoading()
		Click "Up"
		waitLoading()
		Send "{Backspace}"
		waitLoading()
		Send Format("{Text}{1}  {2}", flightIn, tripNumber)
		waitLoading()
		MouseMove 594, 414
		waitLoading()
		Send "!o"
		waitLoading()
		MouseMove 812, 507
		waitLoading()
		Click
		waitLoading()
	}

	static dateTimeEntry(checkin, checkout, ETA, ETD, initX := 323, initY := 506) {
		; fill-in checkin/checkout
		MouseMove 345, initY - 150 ; 332, 356
		waitLoading()
		Click 1
		waitLoading()
		Send "!c"
		waitLoading()
		Send Format("{Text}{1}", checkin)
		waitLoading()
		MouseMove initX + 2, initY - 108 ; 325, 398
		waitLoading()
		Click
		waitLoading()
		MouseMove initX + 338, initY + 37 ; 661, 543
		waitLoading()
		Click
		MouseMove initX + 313, initY + 37 ; 636, 543
		waitLoading()
		Click
		MouseMove initX + 312, initY + 37 ; 635, 543
		waitLoading()
		Click
		waitLoading()
		Click
		waitLoading()
		MouseMove 345, initY - 101 ; 335, 405
		waitLoading()
		Click 1
		waitLoading()
		Send "!c"
		waitLoading()
		Send Format("{Text}{1}", checkout)
		waitLoading()
		Send "{Enter}"
		waitLoading()
		loop 5 {
			Send "{Esc}"
			waitLoading()
		}
		; fill in ETA & ETD
		MouseMove 320, 599
		waitLoading()
		Click 3
		waitLoading()
		Send Format("{Text}{1}", ETA)
		waitLoading()
		Send "{Tab}"
		waitLoading()
		MouseMove 454, 599
		waitLoading()
		Click 3
		waitLoading()
		Send Format("{Text}{1}", ETD)
		Send "{Tab}"
		waitLoading()
	}

	static commentIbdTripEntry(comment, flightIn, tripNumber, initX := 622, initY := 589) {
		; select all and re-enter comment
		MouseMove initX, initY ; 622, 596
		waitLoading()
		Click "Down"
		MouseMove initX + 518, initY + 36 ; 1140, 605
		waitLoading()
		Click "Up"
		waitLoading()
		Send "{Backspace}"
		waitLoading()
		Send comment
		waitLoading()
		; fill-in new flight and trip
		MouseMove initX + 307, initY - 35 ; 929, 554
		waitLoading()
		Click 3
		waitLoading()
		Send Format("{Text}{1}  {2}", flightIn, tripNumber)
		waitLoading()
	}

	static moreFieldsEntry(sCheckin, sCheckout, ETA, ETD, flightIn, flightOut, initX := 236, initY := 333) {
		MouseMove initX, initY ; 236, 333
		waitLoading()
        Click
		waitLoading()
        MouseMove 680, 460
		waitLoading()
        Click 2
		waitLoading()
        Send Format("{Text}{1}", flightIn)
		waitLoading()
        loop 2 {
            Send "{Tab}"
			waitLoading()
        }
        Send Format("{Text}{1}", sCheckin)
        Sleep 100
        Send "{Tab}"
		waitLoading()
        Send Format("{Text}{1}", ETA)
		waitLoading()
        MouseMove 917, 465
		waitLoading()
        Click 2
		waitLoading()
        Send Format("{Text}{1}", flightOut)
		waitLoading()
        loop 2 {
            Send "{Tab}"
			waitLoading()
        }
		waitLoading()
        Send Format("{Text}{1}", sCheckout)
		waitLoading()
        Send "{Tab}"
		waitLoading()
        Send Format("{Text}{1}", ETD)
		waitLoading()
        MouseMove initX + 605, initY + 347 ; 841, 680
		waitLoading()
        Click
		waitLoading()
	}

	static dailyDetailsEntry(daysActual, initX := 372, initY := 524) {
		MouseMove initX, initY ; 372, 524
		waitLoading()
		Click
		waitLoading()
		Send "!d"
		waitLoading()
		loop daysActual {
			Send "{Down}"
			waitLoading()
		}
		Send "!e"
		waitLoading()
		loop 4 {
			Send "{Tab}"
			waitLoading()		
		}
		Send "{Text}NRR"
		waitLoading()
		Send "!o"
		loop 3 {
			Send "{Esc}"
			waitLoading()
		}
		waitLoading()
		Send "!o"
		waitLoading()
		loop 5 {
			Send "{Esc}"
		waitLoading()
		}
		waitLoading()
	}
}