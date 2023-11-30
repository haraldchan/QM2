; #Include "%A_ScriptDir\utils.ahk%"
#Include "./utils.ahk"

FsrMain() {
	WinMaximize "ahk_class SunAwtFrame"
	WinActivate "ahk_class SunAwtFrame"

	sheetIndex := InputBox("
	(
	请输入需要读入的标签页名称。

	如：
	读取"Sheet1"则填入"1"，读取"Sheet1-2"则填入"1-2"
	)", "FedexScheduledReservations")
	if (sheetIndex.Result = "Cancel") {
		cleanReload()
	}
	; schedule reading prep
	row := 4
	path := IniRead(A_ScriptDir . "\config.ini", "FSR", "schedulePath")
	; TODO: check parse "Z:\" to "10.0.2.13\sm\" <- find out how the path is styled
	resvOnDayTime := IniRead(A_ScriptDir . "\config.ini", "FSR", "resvOnDayTime")
	bringForwardTime := getBringForwardTime(resvOnDayTime)
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

	Xl := ComObject("Excel.Application")
	Xlbook := Xl.Workbooks.Open(path)
	shcdDay := Xlbook.Worksheets(Format("Sheet{1}", sheetIndex.Value)) 
	lastRow := Xlbook.ActiveSheet.Cells(Xlbook.ActiveSheet.Rows.Count, "A").End(-4162).Row
	; MsgBox(lastRow)

	; filling in pmsreservations
	BlockInput true
	WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
	loop (lastRow - 3) {
		; receiving shedule info (row)
		flightInfo := Map()
		loop flightFormat.Length {
			flightInfo[flightFormat[A_Index]] := shcdDay.Cells(row, A_Index).Text
		}
		; date reformatting
		myYear := (StrSplit(flightInfo["ibDate"], "/")[1] < A_MM) 
			? A_Year + 1 
			: A_Year
		; schdCi/CoDate format : yyyyMMdd
		schdCiDate := Format("{1}{2}{3}", myYear, StrSplit(flightInfo["ibDate"], "/")[1], StrSplit(flightInfo["ibDate"], "/")[2])
		schdCoDate := Format("{1}{2}{3}", myYear, StrSplit(flightInfo["obDate"], "/")[1], StrSplit(flightInfo["obDate"], "/")[2])
		daysActual := getDaysActual(flightInfo["stayHours"])

		pmsCiDate := (StrSplit(flightInfo["ETA"], ":")[1]) < bringForwardTime 
			? DateAdd(schdCiDate, -1, "days")
			: schdCiDate
		pmsCoDate := schdCoDate
		pmsNts := DateDiff(pmsCoDate, pmsCiDate, "days")
		comment := Format("RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}days, ActualStay: {3}-{4}", flightInfo["stayHours"], daysActual, schdCiDate, schdCoDate)
		; reformat to match pms date format
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
			Sleep 2000
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
			Sleep 300
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
			Sleep 2000
			; { fill in ETA & ETD
			MouseMove 294, 577
			Sleep 200
			Click
			Sleep 200
			Send "{Enter}"
			Sleep 200
			Send "{Enter}"
			Sleep 200
			Send "{Enter}"
			Sleep 200
			MouseMove 320, 577
			Sleep 200
			Click "Down"
			MouseMove 200, 577
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", flightInfo["ETA"])
			Sleep 200
			MouseMove 499, 577
			Sleep 200
			Click "Down"
			MouseMove 330, 574
			Sleep 200
			Click "Up"
			Sleep 200
			Send Format("{Text}{1}", flightInfo["ETD"])
			Sleep 200
			; }
			Sleep 200
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
			Sleep 2000
			; }
			Sleep 100
			; { fill in original shcedule info in "More Fields"
			MouseMove 236, 313
			Sleep 100
			Click
			Sleep 100
			MouseMove 686, 439
			Sleep 200
			Click "Down"
			MouseMove 478, 439
			Sleep 100
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
			Sleep 100
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
			Sleep 2000
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
				loop 4 {
					Send "{Tab}"
					Sleep 100
				}
				Send "{Text}NRR"
				Sleep 100
				MouseMove 418, 377
				Sleep 100
				Send "!o"
				Sleep 200
				loop 3 {
					Send "{Escape}"
					Sleep 200
				}
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
			Sleep 2000
			; { close, save, down to the next one
			Send "!o"
			Sleep 5000
			Send "!o"
			Sleep 1000
			Send "{Down}"
			Sleep 2000
			; }
		}
		row++
	}
	BlockInput false
	WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
	Xlbook.Close()
	Xl.Quit()
	Sleep 1000
	MsgBox("已完成FedEx 预订录入，请抽检以确保准确！", "FedexScheduledReservations")
	cleanReload()
}