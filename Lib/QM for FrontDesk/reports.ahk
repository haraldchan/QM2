; now runs in CoordMode "Screen"
reportOpen(searchStr, initX := 591, initY := 456) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove initX, initY  ; 591, 456
    Sleep 150
    Send "!m"
    Sleep 300
    Send "{Text}R"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    ; print to PDF
    MouseMove initX + 194, initY - 225  ; 785, 231
    Sleep 150
    Click
    MouseMove initX - 104, initY + 133  ; 487, 589
    Sleep 150
    Click
    MouseMove initX + 14, initY - 123  ; 605, 333
    Sleep 150
    Click
    Sleep 50
    Click
    Sleep 150
    BlockInput false
}

reportOpenDelimitedData(searchStr, initX := 591, initY := 456) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove initX, initY
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove initX + 194, initY - 225  ; 785, 231
    Sleep 150
    Click
    MouseMove initX - 135, initY + 145 ; 456, 601
    Sleep 150
    Click "Down"
    MouseMove initX - 132, initY + 145 ; 459, 601
    Sleep 100
    Click "Up"
    MouseMove initX + 225, initY + 142 ; 816, 598
    Sleep 150
    Click
    MouseMove initX + 182, initY + 246 ; 773, 702
    Sleep 150
    Click
    MouseMove initX + 145, initY + 233 ; 736, 689
    Sleep 150
    Click
    BlockInput false
}

reportOpenSpreadSheet(searchStr, initX := 591, initY := 456) {
    WinMaximize "ahk_class SunAwtFrame"
    WinActivate "ahk_class SunAwtFrame"
    BlockInput true
    MouseMove initX, initY ; 591, 456
    Sleep 150
    Send "!m"
    Sleep 150
    Send "R"
    Sleep 150
    Sleep 200
    Send Format("{Text}{1}", searchStr)
    Sleep 200
    MouseMove initX + 206, initY - 221 ; 797, 235
    Sleep 100
    Click
    MouseMove initX - 136, initY + 147 ; 455, 603
    Sleep 100
    Click
    MouseMove initX + 137, initY + 143 ; 728, 599
    Sleep 100
    Click
    MouseMove initX + 166, initY + 267 ; 757, 723
    Sleep 100
    Click
    MouseMove initX + 140, initY + 233 ; 731, 689
    Sleep 100
    Click
    BlockInput false
}

reportSave(savename, initX := 600, initY := 620) {
    BlockInput true
    Sleep 1000
    Send "!f"
    Sleep 1000
    Send "{Backspace}"
    Sleep 200
    Send Format("{Text}{1}", savename)
    Sleep 1000
    Send "{Enter}"
    saveMsg := Format("{1} 保存中", savename)
    MsgBox(saveMsg, "ReportMaster", "T10 4096")
    if (savename = "VIP INH-Guest INH without due out.pdf") {
        Sleep 13000
    }
    Sleep 200
    MouseMove initX, initY
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

mgrFlash(initX := 749, initY := 333) {
    searchStr := "FI01"
    fileName := "NA02-MANAGER FLASH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 749, 333
    Sleep 150
    Send "!o"
    MouseMove initX - 149, initY + 169 ; 600, 502
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisFor15(initX := 644, initY := 178) {
    searchStr := "RS05"
    fileName := "RS05-林总.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 644, 178
    Sleep 150
    Click "Down"
    MouseMove initX - 102, initY - 8 ; 542, 170
    Sleep 150
    Click "Up"
    MouseMove initX - 117, initY - 27 ; 527, 151
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove initX + 2, initY + 26 ; 646, 204
    Sleep 150
    Click "Down"
    MouseMove initX - 111, initY + 41 ; 533, 219
    Sleep 150
    Click "Up"
    MouseMove initX - 171, initY + 92 ; 473, 270
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "8"
    MouseMove initX - 27, initY + 463 ; 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove initX - 34, initY + 511 ; 610, 689
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisForThisMonth(initX := 645, initY := 205) {
    preAuditDate := DateAdd(A_Now, -1, "Days")
    preAuditMonth := FormatTime(preAuditDate, "MM")
    preAuditYear := FormatTime(preAuditDate, "yyyy")
    nextMonth := (preAuditMonth = 12) ? 1 : preAuditMonth + 1

    nextMonth := (nextMonth < 10) ? "0" . nextMonth : nextMonth

    printYear := (preAuditMonth) = 12 ? preAuditYear + 1 : preAuditYear
    firstDayOfNextMonth := printYear . nextMonth . "01"

    dateFirst := preAuditMonth . "01" . preAuditYear
    dateLast := FormatTime(DateAdd(firstDayOfNextMonth, -1, "Days"), "MMddyyyy")

    searchStr := "RS05"
    fileName := Format("RS05-{1}月.pdf", preAuditMonth)
    BlockInput true
    reportOpen(searchStr)
    ; report options here
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateFirst)
    Sleep 300
    MouseMove initX, initY ; 645, 205
    Sleep 150
    Click "Down"
    MouseMove initX - 114, initY - 1 ; 531, 204
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLast)
    Sleep 300
    MouseMove initX - 28, initY + 436 ; 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove initX - 35, initY + 484 ; 610, 689
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

hisForNextMonth(initX := 645, initY := 205) {
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
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateFirstNext)
    Sleep 300
    MouseMove initX, initY ; 645, 205
    Sleep 150
    Click "Down"
    MouseMove initX - 114, initY - 1 ; 531, 204
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{Backspace}"
    Sleep 300
    Send Format("{Text}{1}", dateLastNext)
    Sleep 300
    MouseMove initX - 28, initY + 436 ; 617, 641
    Sleep 150
    Send "{Tab}"
    MouseMove initX - 35, initY + 484 ; 610, 689
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

vipArr(initX := 464, initY := 194) {
    searchStr := "FO01-VIP"
    fileName := "FO01-VIP ARR.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 464, 194
    Sleep 150
    Click "Down"
    MouseMove initX - 128, initY + 4 ; 336, 198
    Sleep 50
    Click "Up"
    Sleep 150
    Send "{NumpadAdd}"
    Sleep 150
    Send "{Text}2"
    Sleep 100
    MouseMove initX + 146, initY + 426 ; 610, 620
    Sleep 150
    Click
    MouseMove initX + 168, initY + 399 ; 632, 593
    Sleep 150
    Click
    MouseMove initX + 146, initY + 549 ; 610, 743
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

vipDep(initX := 600, initY := 545) {
    searchStr := "FO03"
    fileName := "FO03-VIP DEP.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 600, 545
    Sleep 150
    Click
    MouseMove initX - 32, initY - 158 ; 632, 387
    Sleep 150
    Click
    MouseMove initX + 264, initY - 160 ; 864, 385
    Sleep 150
    Click
    MouseMove initX + 222, initY - 91 ; 822, 454
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove initX + 250, initY + 33 ; 850, 578
    Sleep 150
    Click
    MouseMove initX + 35, initY + 136 ; 635, 681
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

arrAll(initX := 309, initY := 566) {
    searchStr := "FO01"
    fileName := "FO01-Arrival Detailed.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 309, 566
    Sleep 150
    Click
    MouseMove initX + 8, initY + 53 ; 317, 619
    Sleep 150
    Click
    MouseMove initX + 16, initY + 97 ; 325, 663
    Sleep 150
    Click
    MouseMove initX + 290, initY + 28 ; 599, 594
    Sleep 150
    Click
    MouseMove initX + 299, initY + 52 ; 608, 618
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

inhAll(initX := 433, initY := 523) {
    searchStr := "FO02"
    fileName := "FO02-INH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 433, 523
    Sleep 150
    Click
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

depAll(initX := 607, initY := 540) {
    searchStr := "FO03"
    fileName := "FO03-DEP.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 607, 540
    Sleep 150
    click
    Sleep 150
    MouseMove initX - 119, initY - 50 ; 488, 490
    Sleep 150
    Click
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

creditLimit(initX := 686, initY := 474) {
    searchStr := "FO11"
    fileName := "FO11-CREDIT LIMIT.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 686, 474
    Click
    MouseMove initX - 277, initY + 48 ; 409, 522
    Sleep 150
    Click
    Sleep 200
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

bbf(initX := 599, initY := 276) {
    searchStr := "FO13"
    fileName := "FO13-Packages 早餐.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 599, 276
    Sleep 150
    Click "Down"
    MouseMove initX - 80, initY + 3 ; 519, 279
    Sleep 150
    Click "Up"
    Sleep 150
    Send "{NumpadSub}"
    Sleep 150
    Send "1"
    Sleep 150
    Send "{Enter}"
    MouseMove initX + 10, initY - 7 ; 609, 269
    Sleep 150
    Click "Down"
    MouseMove initX - 129, initY + 7 ; 470, 283
    Sleep 150
    Click "Up"
    Sleep 100
    Send "^c"
    Sleep 200
    MouseMove initX + 6, initY + 27 ; 605, 303
    Sleep 150
    MouseMove initX + 6, initY + 27 ; 605, 303
    Sleep 150
    Click "Down"
    MouseMove initX - 103, initY + 30 ; 496, 306
    Sleep 150
    Click "Up"
    MouseMove initX - 105, initY + 31 ; 494, 307
    Sleep 100
    Send "^v"
    Sleep 100
    MouseMove initX - 136, initY + 242 ; 463, 518
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

rooms(initX := 476, initY := 515) {
    searchStr := "%hkroomstatusperroom"
    fileName := "Rooms.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY
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

groupRoom(initX := 372, initY := 544) {
    searchStr := "GRPRMLIST"
    fileName := "Group Rooming List.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 372, 544
    Sleep 150
    Click
    MouseMove initX + 148, initY - 78 ; 520, 466
    Sleep 150
    Click
    MouseMove initX + 160, initY + 5 ; 532, 549
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

groupInh(initX := 470, initY := 425) {
    searchStr := "GROUP IN HOUSE"
    fileName := "Group INH.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 470, 425
    Sleep 150
    Click
    MouseMove initX + 155, initY + 68 ; 625, 493
    Sleep 150
    Click
    Sleep 150
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

noShow(initX := 573, initY := 440) {
    searchStr := "FO08"
    fileName := "FO08-NO SHOW.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY
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

cancel(initX := 601, initY := 291) {
    searchStr := "RESCANCEL"
    fileName := "CXL.pdf"
    BlockInput true
    reportOpen(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 601, 291
    Sleep 150
    Click "Down"
    MouseMove initX - 196, initY - 13 ; 485, 278
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
    MouseMove initX + 1, initY - 3 ; 602, 288
    Sleep 150
    Click "Down"
    MouseMove initX - 153, initY + 1 ; 448, 292
    Sleep 150
    Click "Up"
    Sleep 50
    Send "^c"
    Sleep 150
    MouseMove initX - 15, initY + 18 ; 586, 309
    Sleep 150
    Click "Down"
    MouseMove initX - 188, initY + 30 ; 413, 321
    Sleep 150
    Click "Up"
    MouseMove initX - 162, initY + 34 ; 439, 325
    Sleep 50
    Send A_Clipboard
    MouseMove initX + 86, initY + 26 ; 687, 317
    Sleep 150
    Click
    MouseMove initX + 94, initY + 32 ; 695, 323
    Sleep 150

    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

arrivingGroups(blockCodeInput, saveName, initX := 845, initY := 376) {
    fileName := Format("{1}.pdf", saveName)
    BlockInput true
    reportOpen("GRPRMLIST")
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 845, 376
    Sleep 150
    Click
    MouseMove initX - 41, initY + 10 ; 804, 386
    Sleep 150
    Send "!a"
    Sleep 150
    MouseMove initX - 37, initY + 229 ; 808, 605
    Sleep 150
    Click "Down"
    MouseMove initX - 40, initY + 239 ; 805, 614
    Sleep 150
    Click "Up"
    MouseMove initX - 108, initY + 239 ; 737, 615
    Sleep 150
    Click
    Sleep 150
    loop 8 {
        Send "{Space}"
        Sleep 100
        Send "{Up}"
        Sleep 100
    }
    Sleep 350
    Send "!o"
    Sleep 150
    MouseMove initX - 3, initY - 52 ; 842, 324
    Sleep 150
    Click
    MouseMove initX - 276, initY - 93 ; 569, 283
    Sleep 150
    Click
    Sleep 150
    Send Format("{Text}{1}", blockCodeInput)
    Sleep 150
    MouseMove initX + 3, initY - 95 ; 848, 281
    Sleep 150
    Click
    MouseMove initX - 242, initY + 36 ; 603, 412
    Sleep 150
    Click
    Sleep 150
    Send "{Space}"
    MouseMove initX - 117, initY + 97 ; 728, 473
    Sleep 150
    Send "!o"
    Sleep 150
    Send "!o"
    Sleep 150
    MouseMove initX - 134, initY + 115 ; 711, 491
    Sleep 150
    Click
    MouseMove initX - 423, initY + 144 ; 422, 520
    Sleep 150
    Click
    MouseMove initX - 236, initY + 271 ; 609, 647
    reportSave(fileName)
    TrayTip Format("正在保存：{1}", fileName)
    BlockInput false
}

special(today, initX := 600, initY := 482) {
    searchStr := "Wshgz_special"
    fileName := Format("{1} 水果5.xls", today)
    BlockInput true
    reportOpenSpreadSheet(searchStr)
    Sleep 200
    ; report options here
    MouseMove initX, initY ; 600, 482
    Sleep 200
    Click
    Sleep 200
    Send "{Text}水果5"
    Sleep 100
    MouseMove initX + 16, initY + 56 ; 616, 538
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
