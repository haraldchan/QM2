; now runs in CoordMode "Screen"
reportOpen(searchStr) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove 591, 456
    Sleep 150
    Send "!m"
    Sleep 300
    Send "{Text}R"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    ; print to PDF
    MouseMove 785, 231
    Sleep 150
    Click
    MouseMove 487, 589
    Sleep 150
    Click
    MouseMove 605, 333
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
    MouseMove 591, 456
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove 786, 231
    Sleep 150
    Click
    MouseMove 456, 601
    Sleep 150
    Click "Down"
    MouseMove 459, 601
    Sleep 100
    Click "Up"
    MouseMove 816, 598
    Sleep 150
    Click
    MouseMove 773, 702
    Sleep 150
    Click
    MouseMove 736, 689
    Sleep 150
    Click
    BlockInput false
}

reportOpenSpreadSheet(searchStr) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove 591, 456
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 150
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove 797, 235
    Sleep 100
    Click
    MouseMove 455, 603
    Sleep 100
    Click
    MouseMove 728, 599
    Sleep 100
    Click
    MouseMove 757, 723
    Sleep 100
    Click
    MouseMove 731, 689
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
    MsgBox(saveMsg, "ReportMaster", "T10 4096")
    if (savename = "VIP INH-Guest INH without due out.pdf") {
        Sleep 13000
    }
    Sleep 200
    MouseMove 600, 620
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
    MouseMove 749, 333
    Sleep 150
    Send "!o"
    MouseMove 600, 502
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
    MouseMove 644, 178
    Sleep 150
    Click "Down"
    MouseMove 542, 170
    Sleep 150
    Click "Up"
    MouseMove 527, 151
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove 646, 204
    Sleep 150
    Click "Down"
    MouseMove 533, 219
    Sleep 150
    Click "Up"
    MouseMove 473, 270
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 689
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
    preAuditYear := FormatTime(preAuditDate, "yyyy")
    nextMonth := (preAuditMonth = 12) ? 1 : preAuditMonth + 1
    ; if (nextMonth < 10) { 
    ;     nextMonth := Format("0{1}", nextMonth)
    ; }
    nextMonth := (nextMonth < 10) ? "0" . nextMonth : nextMonth

    printYear := (preAuditMonth) = 12 ? preAuditYear + 1 : preAuditYear
    firstDayOfNextMonth := printYear . nextMonth . "01"

    dateFirst := preAuditMonth . "01" . preAuditYear
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
    MouseMove 645, 205
    Sleep 150
    Click "Down"
    MouseMove 531, 204
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLast)
    Sleep 300
    MouseMove 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 689
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
    preAuditYear := FormatTime(preAuditDate, "yyyy")
    nMonth := (preAuditMonth = 12) ? 1 : preAuditMonth + 1
    nextMonth := (nMonth < 10) ? "0" . nMonth : nMonth
    nextNextMonth := (nextMonth = 12) ? 1 : nextMonth + 1
    nextNextMonth := (nextNextMonth < 10) ? "0" . nextNextMonth : nextNextMonth
    printYear := preAuditMonth = 12 ? preAuditYear + 1 : preAuditYear
    firstDayOfNextMonth := printYear . nextMonth . "01"
    firstDayOfNextNextMonth := printYear . nextNextMonth . "01"
    dateFirstNext := nextMonth . "01" . printYear
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
    MouseMove 645, 205
    Sleep 150
    Click "Down"
    MouseMove 531, 204
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLastNext)
    Sleep 300
    MouseMove 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove 610, 689
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
    MouseMove 464, 194
    Sleep 150
    Click "Down"
    MouseMove 336, 198
    Sleep 50
    Click "Up"
    Sleep 800
    Send "{NumpadAdd}"
    Sleep 150
    Send "{Text}2"
    Sleep 100
    MouseMove 610, 620
    Sleep 150
    Click
    MouseMove 632, 593
    Sleep 150
    Click
    MouseMove 610, 743
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
    MouseMove 600, 545
    Sleep 150
    Click
    MouseMove 632, 387
    Sleep 150
    Click
    MouseMove 864, 385
    Sleep 150
    Click
    MouseMove 822, 454
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove 850, 578
    Sleep 150
    Click
    MouseMove 635, 681
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
    MouseMove 309, 566
    Sleep 150
    Click
    MouseMove 317, 619
    Sleep 150
    Click
    MouseMove 325, 663
    Sleep 150
    Click
    MouseMove 599, 594
    Sleep 150
    Click
    MouseMove 608, 618
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
    MouseMove 433, 523
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
    MouseMove 607, 540
    Sleep 150
    click
    Sleep 150
    MouseMove 488, 490
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
    MouseMove 686, 474
    Click
    MouseMove 409, 522
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
    MouseMove 599, 276
    Sleep 150
    Click "Down"
    MouseMove 519, 279
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "1"
    Sleep 150
    Send "{Enter}"
    MouseMove 609, 269
    Sleep 150
    Click "Down"
    MouseMove 470, 283
    Sleep 150
    Click "Up"
    Sleep 100
    Send "^c"
    Sleep 200
    MouseMove 605, 303
    Sleep 150
    MouseMove 605, 303
    Sleep 150
    Click "Down"
    MouseMove 496, 306
    Sleep 150
    Click "Up"
    MouseMove 494, 307
    Sleep 100
    Send "^v"
    Sleep 100
    MouseMove 463, 518
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
    MouseMove 476, 515
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
    MouseMove 372, 544
    Sleep 150
    Click
    MouseMove 520, 466
    Sleep 150
    Click
    MouseMove 532, 549
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
    MouseMove 470, 425
    Sleep 150
    Click
    MouseMove 625, 493
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
    MouseMove 573, 440
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
    MouseMove 601, 291
    Sleep 150
    Click "Down"
    MouseMove 485, 278
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
    MouseMove 602, 288
    Sleep 150
    Click "Down"
    MouseMove 448, 292
    Sleep 150
    Click "Up"
    Sleep 50
    Send "^c"
    Sleep 150
    MouseMove 586, 309
    Sleep 150
    Click "Down"
    MouseMove 413, 321
    Sleep 150
    Click "Up"
    MouseMove 439, 325
    Sleep 50
    Send A_Clipboard
    MouseMove 687, 317
    Sleep 150
    Click
    MouseMove 695, 323
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
    MouseMove 845, 376
    Sleep 150
    Click
    MouseMove 804, 386
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove 808, 605
    Sleep 150
    Click "Down"
    MouseMove 805, 614
    Sleep 150
    Click "Up"
    MouseMove 737, 615
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
    MouseMove 842, 324
    Sleep 150
    Click
    MouseMove 569, 283
    Sleep 150
    Click
    Sleep 150
    Send Format("{Text}{1}", blockCodeInput)
    Sleep 150
    MouseMove 848, 281
    Sleep 150
    Click
    MouseMove 603, 412
    Sleep 150
    Click
    Sleep 150
    Send "{Space}"
    MouseMove 728, 473
    Sleep 150
    Send "!o"
    Sleep 150
    Send "!o"
    Sleep 150
    MouseMove 711, 491
    Sleep 150
    Click
    MouseMove 422, 520
    Sleep 150
    Click
    MouseMove 609, 647
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
    MouseMove 600, 482
    Sleep 200
    Click
    Sleep 200
    Send "{Text}水果5"
    Sleep 100
    MouseMove 616, 538
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
