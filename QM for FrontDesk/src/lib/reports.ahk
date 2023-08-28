reportOpen(searchStr) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove 591, 436
    Sleep 150
    Send "!m"
    Sleep 300
    Send "{R}"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    ; print to PDF
    MouseMove 785, 211
    Sleep 150
    Click
    MouseMove 487, 569
    Sleep 150
    Click
    MouseMove 605, 313
    Sleep 150
    Click
    Sleep 50
    Click
    Sleep 150
    BlockInput false
}

reportOpenDelimitedData(searchStr) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove 591, 436
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove 756, 222
    MouseMove 786, 211
    Sleep 150
    Click
    MouseMove 456, 581
    Sleep 150
    Click "Down"
    MouseMove 459, 581
    Sleep 100
    Click "Up"
    MouseMove 816, 578
    Sleep 150
    Click
    MouseMove 773, 682
    Sleep 150
    Click
    MouseMove 736, 669
    Sleep 150
    Click
    BlockInput false
}

reportOpenSpreadSheet(searchStr) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove 591, 436
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 150
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove 797, 215
    Sleep 100
    Click
    MouseMove 455, 583
    Sleep 100
    Click
    MouseMove 728, 579
    Sleep 100
    Click
    MouseMove 757, 703
    Sleep 100
    Click
    MouseMove 731, 669
    Sleep 100
    Click
    BlockInput false
}
reportSave(savename) {
    BlockInput true
    Sleep 1000
    Send "!f"
    Sleep 1000
    Send "{Backspace}"
    Sleep 1000
    Send Format("{Text}{1}", savename)
    Sleep 200
    Send "{Enter}"
    saveMsg := Format("{1} 保存中", savename)
    MsgBox(saveMsg, "ReportMaster", "T20 4096")
    if (savename = "VIP INH-Guest INH without due out.pdf") {
        Sleep 13000
    }
    Sleep 200
    MouseMove 600, 600
    Click
    Sleep 200
    Send "!c"
    BlockInput false
}

comp() {
    searchStr := "%complimentary"
    fileName := "comp.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

mgrFlash() {
    searchStr := "FI01"
    fileName := "NA02-MANAGER FLASH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 749, 313
    Sleep 150
    Send "!o"
    MouseMove 600, 482
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisFor15() {
    searchStr := "RS05"
    fileName := "RS05-林总.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 644, 158
    Sleep 150
    Click "Down"
    MouseMove 542, 150
    Sleep 150
    Click "Up"
    MouseMove 527, 131
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove 646, 184
    Sleep 150
    Click "Down"
    MouseMove 533, 199
    Sleep 150
    Click "Up"
    MouseMove 473, 250
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove 617, 621
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 669
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisForThisMonth() {
    preAuditDate := DateAdd(A_Now, -1, "Days")
    preAuditMonth := FormatTime(preAuditDate, "MM")
    preAuditDay := FormatTime(preAuditDate, "dd")
    preAuditYear := FormatTime(preAuditDate, "yyyy")

    nextMonth := preAuditMonth = 12 ? 1 : preAuditMonth + 1
    if (nextMonth < 10) nextMonth := Format("0{1}", nextMonth)
        printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
    firstDayOfNextMonth := Format("{2}{1}01", nextMonth, printYear)

    dateFirst := Format("{1}01{2}", preAuditMonth, preAuditYear)
    dateLast := FormatTime(DateAdd(firstDayOfNextMonth, -1, "Days"), "MMddyyyy")

    searchStr := "RS05"
    fileName := Format("RS05-{1}月.pdf", preAuditMonth)
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateFirst)
    Sleep 300
    MouseMove 645, 185
    Sleep 150
    Click "Down"
    MouseMove 531, 184
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLast)
    Sleep 300
    MouseMove 617, 621
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 669
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisForNextMonth() {
    preAuditDate := DateAdd(A_Now, -1, "Days")
    preAuditMonth := FormatTime(preAuditDate, "MM")
    preAuditDay := FormatTime(preAuditDate, "dd")
    preAuditYear := FormatTime(preAuditDate, "yyyy")

    nMonth := preAuditMonth = 12 ? 1 : preAuditMonth + 1
    if (nMonth < 10) nextMonth := Format("0{1}", nMonth)
        nextNextMonth := nextMonth = 12 ? 1 : nextMonth + 1
    if (nextNextMonth < 10) {
        nextNextMonth := Format("0{1}", nextNextMonth)
    }
    printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
    firstDayOfNextMonth := Format("{2}{1}01", nextMonth, printYear)
    firstDayOfNextNextMonth := Format("{2}{1}01", nextNextMonth, printYear)
    dateFirstNext := Format("{1}01{2}", nextMonth, printYear)
    dateLastNext := FormatTime(DateAdd(firstDayOfNextNextMonth, -1, "Days"), "MMddyyyy")

    searchStr := "RS05"
    fileName := Format("RS05-{1}月.pdf", nextMonth)
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateFirstNext)
    Sleep 300
    MouseMove 645, 185
    Sleep 150
    Click "Down"
    MouseMove 531, 184
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLastNext)
    Sleep 300
    MouseMove 617, 621
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 669
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

vipArr() {
    searchStr := "FO01-VIP"
    fileName := "FO01-VIP ARR.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 464, 174
    Sleep 150
    Click "Down"
    MouseMove 336, 178
    Sleep 50
    Click "Up"
    Sleep 800
    Send "{NumpadAdd}"
    Sleep 150
    Send "{Text}2"
    Sleep 100
    MouseMove 610, 600
    Sleep 150
    Click
    MouseMove 632, 573
    Sleep 150
    Click
    MouseMove 610, 723
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

vipInh() {
    searchStr := "%GUEST INH WITHOUT DUE OUT"
    fileName := "VIP INH-Guest INH without due out.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

vipDep() {
    searchStr := "FO03"
    fileName := "FO03-VIP DEP.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 600, 525
    Sleep 150
    Click
    MouseMove 632, 367
    Sleep 150
    Click
    MouseMove 864, 365
    Sleep 150
    Click
    MouseMove 822, 434
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove 850, 558
    Sleep 150
    Click
    MouseMove 635, 661
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

arrAll() {
    searchStr := "FO01"
    fileName := "FO01-Arrival Detailed.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 309, 546
    Sleep 150
    Click
    MouseMove 317, 599
    Sleep 150
    Click
    MouseMove 325, 643
    Sleep 150
    Click
    MouseMove 599, 574
    Sleep 150
    Click
    MouseMove 608, 598
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

inhAll() {
    searchStr := "FO02"
    fileName := "FO02-INH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 433, 503
    Sleep 150
    Click
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

depAll() {
    searchStr := "FO03"
    fileName := "FO03-DEP.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 607, 520
    Sleep 150
    click
    Sleep 150
    MouseMove 488, 470
    Sleep 150
    Click
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

creditLimit() {
    searchStr := "FO11"
    fileName := "FO11-CREDIT LIMIT.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 686, 454
    Click
    MouseMove 409, 502
    Sleep 150
    Click
    Sleep 200

    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

bbf() {
    searchStr := "FO13"
    fileName := "FO13-Packages 早餐.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 599, 256
    Sleep 150
    Click "Down"
    MouseMove 519, 259
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "1"
    Sleep 150
    Send "{Enter}"
    MouseMove 609, 249
    Sleep 150
    Click "Down"
    MouseMove 470, 263
    Sleep 150
    Click "Up"
    Sleep 100
    Send "^c"
    Sleep 200
    MouseMove 605, 283
    Sleep 150
    MouseMove 605, 283
    Sleep 150
    Click "Down"
    MouseMove 496, 286
    Sleep 150
    Click "Up"
    MouseMove 494, 287
    Sleep 100
    Send "^v"
    Sleep 100
    MouseMove 463, 498
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

rooms() {
    searchStr := "%hkroomstatusperroom"
    fileName := "Rooms.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 476, 495
    Sleep 150
    Click
    Sleep 150

    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

ooo() {
    searchStr := "HK03"
    fileName := "HK03-OOO.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

groupRoom() {
    searchStr := "GRPRMLIST"
    fileName := "Group Rooming List.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 372, 524
    Sleep 150
    Click
    MouseMove 520, 446
    Sleep 150
    Click
    MouseMove 532, 529
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

groupInh() {
    searchStr := "GROUP IN HOUSE"
    fileName := "Group INH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 470, 405
    Sleep 150
    Click
    MouseMove 625, 473
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

noShow() {
    searchStr := "FO08"
    fileName := "FO08-NO SHOW.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 573, 420
    Sleep 50
    Send "{Backspace}"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 100
    Send "{Text}1"
    Sleep 150
    Send "{Tab}"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 100
    Send "{Text}1"
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

cancel() {
    searchStr := "RESCANCEL"
    fileName := "CXL.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove 601, 271
    Sleep 150
    Click "Down"
    MouseMove 485, 258
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "{Text}1"
    Sleep 50
    Send "{Enter}"
    Sleep 150
    MouseMove 602, 268
    Sleep 150
    Click "Down"
    MouseMove 448, 272
    Sleep 150
    Click "Up"
    Sleep 50
    Send "^c"
    Sleep 150
    MouseMove 586, 289
    Sleep 150
    Click "Down"
    MouseMove 413, 301
    Sleep 150
    Click "Up"
    MouseMove 439, 305
    Sleep 50
    Send A_Clipboard
    MouseMove 687, 297
    Sleep 150
    Click
    MouseMove 695, 303
    Sleep 150

    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

arrivingGroups(blockCodeInput, saveName) {
    fileName := Format("{1}.pdf", saveName)
    BlockInput true
    reportOpen("GRPRMLIST")
    Sleep 200
    ; report options here
    MouseMove 845, 356
    Sleep 150
    Click
    MouseMove 804, 366
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove 808, 585
    Sleep 150
    Click "Down"
    MouseMove 805, 594
    Sleep 150
    Click "Up"
    MouseMove 737, 595
    Sleep 150
    Click
    Sleep 150
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 100
    Send "{Space}"
    Sleep 100
    Send "{Up}"
    Sleep 350
    Send "!o"
    Sleep 150
    MouseMove 842, 304
    Sleep 150
    Click
    MouseMove 569, 263
    Sleep 150
    Click
    Sleep 150
    Send Format("{Text}{1}", blockCodeInput)
    Sleep 150
    MouseMove 848, 261
    Sleep 150
    Click
    MouseMove 603, 392
    Sleep 150
    Click
    Sleep 150
    Send "{Space}"
    MouseMove 728, 453
    Sleep 150
    Send "!o"
    Sleep 150
    Send "!o"
    Sleep 150
    MouseMove 711, 471
    Sleep 150
    Click
    MouseMove 422, 500
    Sleep 150
    Click
    MouseMove 609, 627
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

special(today) {
    searchStr := "Wshgz_special"
    fileName := Format("{1} 水果5.xls", today)
    BlockInput true
    reportOpenSpreadSheet(searchStr)
    Sleep 200
    ; report options here
    MouseMove 600, 462
    Sleep 200
    Click
    Sleep 200
    Send "{Text}水果5"
    Sleep 100
    MouseMove 616, 518
    Sleep 200
    reportSave(fileName)
    TrayTip Format("正在保存：", fileName)
    BlockInput false
}


; template(){
;     searchStr := ""
;     fileName := ""
;     BlockInput true
;     reportOpen(searchStr)
;     Sleep 200

;     ; report options here

;     reportSave(fileName)
;     TrayTip Format("正在保存：{1}",fileName)
;     BlockInput false
; }


; ^F9::reportOpenDelimitedData("grpr")
; F12:: Reload
; ^F12:: ExitApp
