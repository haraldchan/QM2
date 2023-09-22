; #Include "%A_ScriptDir\utils.ahk%"
#Include "../utils.ahk"

FsrMain() {
	sheetIndex := InputBox("请输入需要读入第几个标签", "FedexScheduledReservations")
	if (sheetIndex.Result = "Cancel") {
		cleanReload()
	}
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	; schedule reading prep
	row := 4
	path := IniRead(A_ScriptDir . "\config.ini", "FSR", "schedulePath")
	resvOnDayTime := IniRead(A_ScriptDir . "\config.ini", "FSR", "toNextDayTime")
	bringForwardTime := getBringForwardTime(resvOnDayTime)
	Xl := ComObject("Excel.Application")
	Xlbook := Xl.Workbooks.Open(path)
	shcdDay := Xlbook.Worksheets("Sheet%sheetIndex.Value%")
	lastRow := Xlbook.ActiveSheet.Cells(Xlbook.ActiveSheet.Rows.Count, "A").End(-4162).Row

	; filling in pmsreservations
	BlockInput true
	loop (lastRow - 4) {
		; receiving shedule info (row)
		flightInfo := Map()
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
		loop flightFormat.Legnth {
			flightInfo[flightFormat[A_Index]] := shcdDay.Cells(row, A_Index).Text
		}
		; date reformatting
		if (StrSplit(flightInfo["ibDate"], "/")[1] < A_MM) {
			myYear := A_Year + 1
		} else {
			myYear := A_Year
		}
		; schdCi/CoDate format : yyyyMMdd
		schdCiDate := Format("{1}{2}{3}", myYear, StrSplit(flightInfo["ibDate"], "/")[1], StrSplit(flightInfo["ibDate"], "/")[2])
		schdCoDate := Format("{1}{2}{3}", myYear, StrSplit(flightInfo["obDate"], "/")[1], StrSplit(flightInfo["obDate"], "/")[2])
		splitHours := StrSplit(flightInfo["stayHours"], ":")
		daysActual := getDaysActual(splitHours[1], splitHours[2])
		if (StrSplit(flightInfo["ETA"], ":")[1]) < bringForwardTime {
			pmsCiDate := DateAdd(schdCiDate, -1, "days")
		} else {
			pmsCiDate := schdCiDate
		}
		pmsCoDate := schdCoDate
		pmsNts := DateDiff(pmsCiDate, pmsCoDate, "days")
		comment := Format("RM INCL 1BBF TO CO,Hours@Hotel:{1}={2}days, ActualStay:{3}-{4}", flightInfo["stayHours"], pmsNts, schdCiDate, schdCoDate)
		; reformat pms to match pms date format
		schdCiDate := FormatTime(schdCiDate, "MMddyyyy")
		schdCoDate := FormatTime(schdCoDate, "MMddyyyy")
		pmsCiDate := FormatTime(pmsCiDate, "MMddyyyy")
		pmsCoDate := FormatTime(pmsCoDate, "MMddyyyy")

		; fill in pms reservations
		loop flightInfo["roomQty"] {
			WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
			; { opening a booking
			Sleep 1000
			MouseMove 399, 244
			Sleep 150
			Send "!e"
			Sleep 1500
			; }
			Sleep 100
			; { opening Profile, fill in inbound flight info & trip no.
			MouseMove 467, 201
			Sleep 300
			Click
			Sleep 1000
			MouseMove 442, 264
			Sleep 300
			Click "Down"
			MouseMove 214, 269
			Sleep 300
			Click "Up"
			Sleep 300
			Send "{Backspace}"
			Sleep 300
			Send Format("{Text}{1}{2}  {3}", flightInfo["flightIn1"], flightInfo["flightIn2"], flightInfo["tripNum"])
			Sleep 300
			MouseMove 594, 394
			Sleep 150
			Send "!c"
			Sleep 300
			MouseMove 576, 521
			Sleep 150
			Click
			Sleep 300
			MouseMove 798, 480
			Sleep 500
			Click
			Sleep 500
			; }
			Sleep 100
			; { fill in check-in, check-out fields
			MouseMove 332, 336
			Sleep 1000
			Click "Down"
			MouseMove 178, 340
			Sleep 300
			Click "Up"
			MouseMove 172, 340
			Sleep 300
			Send Format("{Text}{1}", pmsCiDate)
			Sleep 100
			MouseMove 325, 378
			Sleep 300
			Click
			Sleep 300
			MouseMove 661, 523
			Sleep 300
			Click
			MouseMove 636, 523
			Sleep 300
			Click
			MouseMove 635, 523
			Sleep 300
			Click
			Sleep 300
			Click
			Sleep 2000
			MouseMove 335, 385
			Sleep 300
			Click "Down"
			MouseMove 182, 389
			Sleep 300
			Click "Up"
			MouseMove 207, 395
			Sleep 300
			Send Format("{Text}{1}", pmsCoDate)
			Sleep 300
			; }
			Sleep 100
			; { fill in ETA & ETD
			MouseMove 294, 577
			Sleep 200
			Click
			Sleep 200
			Send "{Enter}"
			Sleep 200
			Send "{Enter}"
			Sleep 200
			MouseMove 332, 577
			Sleep 200
			Click "Down"
			MouseMove 222, 574
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", flightInfo["ETA"])
			Sleep 200
			MouseMove 499, 577
			Sleep 200
			Click "Down"
			MouseMove 342, 574
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", flightInfo["ETD"])
			Sleep 200
			; }
			Sleep 100
			; { fill in comments, Inbound & trip no.(TA REC log field)
			MouseMove 622, 576
			Sleep 200
			Click "Down"
			MouseMove 1140, 585
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", comment)
			Sleep 200
			MouseMove 839, 535
			Sleep 100
			Click "Down"
			MouseMove 1107, 543
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}{2}  {3}", flightInfo["flightIn1"], flightInfo["flightIn2"], flightInfo["tripNum"])
			Sleep 1000
			; }
			Sleep 100
			; { fill in original shcedule info in "More Fields"
			MouseMove 236, 313
			Sleep 100
			Click
			Sleep 100
			MouseMove 686, 439
			Sleep 120
			Click "Down"
			MouseMove 478, 439
			Sleep 342
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}{2}", flightInfo["flightIn1"], flightInfo["flightIn2"])
			Sleep 100
			MouseMove 672, 483
			Sleep 281
			Click "Down"
			MouseMove 523, 483
			Sleep 358
			Click "Up"
			Send Format("{Text}{1}", schdCiDate)
			Sleep 100
			MouseMove 685, 506
			Sleep 100
			Click "Down"
			MouseMove 422, 501
			Sleep 100
			Click "Up"
			Send Format("{Text}{1}", flightInfo["ETA"])
			Sleep 100
			MouseMove 922, 441
			Sleep 100
			Click "Down"
			MouseMove 704, 439
			Sleep 100
			Click "Up"
			Send Format("{Text}{1}{2}", flightInfo["flightOut1"], flightInfo["flightOut2"])
			Sleep 100
			MouseMove 917, 483
			Sleep 100
			Click "Down"
			MouseMove 637, 487
			Sleep 1000
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}", pmsCoDate)
			Sleep 100
			MouseMove 922, 504
			Sleep 100
			Click "Down"
			MouseMove 640, 503
			Sleep 100
			Click "Up"
			Sleep 100
			Send Format("{Text}{1}", flightInfo["ETD"])
			Sleep 100
			MouseMove 841, 660
			Sleep 100
			Click
			Sleep 300
			; }
			Sleep 100
			; { open "Daily Details", fill in room rates
			MouseMove 372, 504
			Sleep 300
			Click
			Sleep 300
			Send "!d"
			Sleep 300
			loop daysActual {
				Send "{Text}1265"
				Send "{Down}"
				Sleep 200
			}
			if (daysActual != pmsNts) {
				Send "0"
				Sleep 200
				MouseMove 618, 485
				Sleep 200
				Send "!e"
				Sleep 200
				Send "{Tab}"
				Sleep 100
				Send "{Tab}"
				Sleep 100
				Send "{Tab}"
				Sleep 100
				Send "{Tab}"
				Sleep 100
				Send "{Text}NRR"
				Sleep 100
				MouseMove 418, 377
				Sleep 100
				Send "!o"
				Sleep 200
				Send "{Escape}"
				Sleep 200
				Send "{Escape}"
				Sleep 200
				Send "!o"
				Sleep 1500
				Send "{Space}"
				Sleep 200
			}
			MouseMove 728, 548
			Sleep 300
			Send "!o"
			Sleep 300
			MouseMove 542, 453
			Sleep 300
			Click
			MouseMove 644, 523
			Sleep 300
			Click
			Sleep 300
			; }
			Sleep 100
			; { close, save, down to the next one
			MouseMove 795, 466
			Sleep 300
			Send "!o"
			Sleep 300
			Send "!o"
			Sleep 300
			Send "!o"
			Sleep 300
			Send "{Down}"
			Sleep 1000
			; }
			row++
			WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
		}
	}
	BlockInput false
	Xlbook.Close()
	Xl.Quit()
	Sleep 1000
	MsgBox("已完成FedEx 预订修改，请抽检以确保准确！", "FedexScheduledReservations")
	cleanReload()
}