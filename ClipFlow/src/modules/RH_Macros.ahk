; FedEx WIP
RH_Fedex(infoObj) {
    pmsCiDate := (StrSplit(infoObj["ETA"], ":")[1]) < 10
        ? DateAdd(infoObj["ciDate"], -1, "days")
        : infoObj["ciDate"]
    pmsCoDate := infoObj["coDate"]
    pmsNts := DateDiff(pmsCoDate, pmsCiDate, "days")
    ; reformat to match pms date format
    schdCiDate := FormatTime(infoObj["ciDate"], "MMddyyyy")
    schdCoDate := FormatTime(infoObj["ciDate"], "MMddyyyy")
    pmsCiDate := FormatTime(pmsCiDate, "MMddyyyy")
    pmsCoDate := FormatTime(pmsCoDate, "MMddyyyy")


    profileEntry(infoObj) {
        crewName := StrSplit(infoObj["crewNames"][A_Index], " ")
        ;TODO: action: open profile, new profile, fill-in first and last name, then save
    }

    dateTimeEntry(infoObj) {
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
        Send Format("{Text}{1}", infoObj["ETA"])
        Sleep 200
        MouseMove 499, 577
        Sleep 200
        Click "Down"
        MouseMove 330, 574
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", infoObj["ETD"])
        Sleep 200
        ; }
        Sleep 200
    }

    commentEntry(infoObj) {
        ; set comment
        comment := ""
        if (infoObj["resvType"] = "ADD") {
            comment := Format("RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}day(s), ActualStay: {3}-{4}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"])
        } else {
            ; TODO: action: copy current comment
            MouseMove 622, 576
            Sleep 200
            Click "Down"
            MouseMove 1140, 585
            Sleep 200
            Click "Up"
            Sleep 200
            Send "^c"
            prevComment := A_Clipboard
            comment := Format("Changed to {1}={2}day(s), New Stay:{3}-{4} `n Before Update:{5}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"], prevComment)
        }
        ; fill-in comment
        Sleep 100
        MouseMove 622, 576
        Sleep 200
        Click "Down"
        MouseMove 1140, 585
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", comment)
        Sleep 200
        ; fill-in new flight and trip
        Sleep 200
        MouseMove 839, 535
        Sleep 100
        Click "Down"
        MouseMove 1107, 543
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}  {2}", infoObj["flightIn"], infoObj["tripNum"])
    }

    moreFieldsEntry(infoObj) {
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
        Send Format("{Text}{1}", infoObj["flightIn"])
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
        Send Format("{Text}{1}", infoObj["ETA"])
        Sleep 100
        MouseMove 922, 441
        Sleep 100
        Click "Down"
        MouseMove 704, 439
        Sleep 100
        Click "Up"
        Send Format("{Text}{1}", infoObj["flightOut"])
        Sleep 100
        MouseMove 917, 483
        Sleep 100
        Click "Down"
        MouseMove 637, 487
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", schdCoDate)
        Sleep 100
        MouseMove 922, 504
        Sleep 100
        Click "Down"
        MouseMove 640, 503
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", infoObj["ETD"])
        Sleep 100
        MouseMove 841, 660
        Sleep 100
        Click
        Sleep 2000
    }

    dailyDetailsEntry(infoObj) {
        MouseMove 372, 504
        Sleep 300
        Click
        Sleep 300
        Send "!d"
        Sleep 300
        loop infoObj["daysActual"] {
            Send Format("{Text}{1}", infoObj["roomRate"][1])
            Send "{Down}"
            Sleep 200
        }
        if (infoObj["daysActual"] != pmsNts) {
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
    }

    addModification(infoObj) {
        addFrom := InputBox("请指定需要从哪个 Block Add-On? (请输入 BlockCode )", "Reservation Handler", , A_Year . A_MM . A_MDay . "FEDEX")
        if (addFrom.Result = "Cancel") {
            cleanReload()
        }

        loop infoObj["roomQty"] {
            if (A_Index = infoObj["roomQty"]) {
                MsgBox("已完成所有新增。")
                return
            }
            ; TODO: action: adds on from a block pm(by using addFrom.Value), then open it.
            profileEntry(infoObj)
            Sleep 1000
            dateTimeEntry(infoObj)
            Sleep 1000
            commentEntry(infoObj)
            Sleep 1000
            moreFieldsEntry(infoObj)
            Sleep 1000
            dailyDetailsEntry(infoObj)
            Sleep 1000
            ; TODO: action: close and save the reservation.
            Sleep 1000
        }
    }

    changeModification(infoObj) {
        loop infoObj["roomQty"] {
            if (A_Index = infoObj["roomQty"]) {
                MsgBox("已完成所有修改。")
                return
            }
            nofitier := Format("待修改订单数量：{1}, 已完成{2}个。`n`n 请先打开需要修改的 Fedex 预订。", infoObj["roomQty"], A_Index)
            changeFrom := MsgBox(nofitier, "Reservation Handler", "OKCancel 4096")
            if (changeFrom.Result = "Cancel") {
                cleanReload()
            }
            profileEntry(infoObj)
            Sleep 1000
            dateTimeEntry(infoObj)
            Sleep 1000
            commentEntry(infoObj)
            Sleep 1000
            moreFieldsEntry(infoObj)
            Sleep 1000
            dailyDetailsEntry(infoObj)
            Sleep 1000
            ; TODO: action close and save the reservation.
            Sleep 1000

        }
    }

    ; starter
    if (infoObj["resvType"] = "CHANGE") {
        changeModification(infoObj)
    } else {
        addModification(infoObj)
    }
}

; Kingsley WIP
RH_Kingsley(infoObj, addFromConf) {
    roomTypeRef := Map([
        "标准大床房", "SKC",
        "标准双床房", "STC",
        "豪华城景大床房", "DKC",
        "豪华城景双床房", "DTC",
        "豪华江景大床房", "DKR",
        "豪华江景双床房", "DTR"
        "行政豪华城景大床房", "CKC",
        "行政豪华城景双床房", "CTC",
        "行政豪华江景大床房", "CKR",
        "行政豪华江景双床房", "CTR"
        "行政尊贵套房", "CSK"
    ])
    roomType := ""
    for k, v in roomTypeRef {
        if (infoObj["roomType"] = k) {
            roomType := v
        }
    }
    addFrom := addFromConf
    breakfastType := (SubStr(roomType, 1, 1) = "C") ? "CBF" : "BBF"
    breakfastQty := infoObj["bbf"][1]
    comment := (breakfastQty = 0)
        ? "RM TO TA"
        : Format("RM INCL {1}{2} TO TA", breakfastQty, breakfastType)
    pmsGuestNames := []
    loop infoObj["guestNames"] {
        curGuestName := infoObj["guestNames"][A_Index]
        if (RegExMatch(curGuestName, "^[a-zA-Z/]+$") > 0) {
            ; if it only includes English alphabet, push [lastName, firstName]
            pmsGuestNames.Push(StrSplit(curGuestName, "/"))
        } else {
            pmsGuestNames.Push([
                getFullnamePinyin(curGuestName)[1], ; lastName pinyin
                getFullnamePinyin(curGuestName)[2], ; firstName pinyin
                curGuestName, ; hanzi-name
            ])
        }
    }

    ; TODO: add-on new booking from addFrom booking
    ; TODO: open the booking, fill-in profile
    profileEntry(pmsGuestNames[1])
    Sleep 1000
    commonEntries(infoObj, roomType, comment)
    Sleep 1000
    ; TODO: close and save
    ; TODO: if roomQty > 1, split and fill-in other names
}

; Agoda WIP
RH_Agoda(infoObj, addFromConf) {
    roomTypeRef := Map([
        "Standard Room - Queen bed", "SKC",
        "Standard Room - Twin Bed", "STC",
        "Deluxe City View with King Bed", "DKC",
        "Deluxe City View Twin Bed", "DTC",
        "Deluxe River View King Bed", "DKR"
        "Deluxe River View Twin Bed", "DTR",
        "Club Deluxe City View", "CKC",
        "Club Deluxe River View Room", "CKR",
        "Club Premier Suite", "CPK",
    ])

    roomType := ""
    for k, v in roomTypeRef {
        if (infoObj["roomType"] = k) {
            roomType := v
        }
    }

    breakfastType := (roomType = "CKC" || roomType = "CKR") ? "CBF" : "BBF"
    comment := Format("ROOM INCL {1}{2} PAY BY AGODA", infoObj["bbf"][1], breakfastType)

    profileEntry(lastName, firstName) {
        ;TODO: action: open profile, new profile, fill-in first and last name, then save
    }

    loop infoObj["roomQty"] {
        profileEntry(infoObj["guestLastName"], infoObj["guestFirstName"])
        Sleep 1000
        commonEntries(infoObj, roomType, comment)
    }
}

profileEntry(guestName) {
    WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
    CoordMode "Mouse", "Screen"
    BlockInput true
    ; TODO: action: new profile.
    MouseMove 433, 283
    Click 3
    Send Format("{Text}{1}", guestName[1])
    Sleep 10
    MouseMove 393, 307
    Click 3
    Send Format("{Text}{1}", guestName[2])
    Sleep 10
    if (guestName.Length = 3) {
        MouseMove 429, 214
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestName[3])
        Sleep 10
    }
    ; TODO: action: save this profile.
}

; commonEntries WIP
commonEntries(infoObj, roomTypeModded, comment) {
    pmsCiDate := FormatTime(infoObj["ciDate"], "MMddyyyy")
    pmsCoDate := FormatTime(infoObj["coDate"], "MMddyyyy")

    roomQtyEntry(roomQty) {
        ; TODO: fill-in roomQty
    }

    dateTimeEntry() {
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
        Sleep 2000
    }

    commentOrderIdEntry(orderId, comment) {
        ; fill-in comment
        Sleep 100
        MouseMove 622, 576
        Sleep 200
        Click "Down"
        MouseMove 1140, 585
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", comment)
        Sleep 200
        ; fill-in orderId
        Sleep 200
        MouseMove 839, 535
        Sleep 100
        Click "Down"
        MouseMove 1107, 543
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", orderId)
    }

    roomRatesEntry(roomRates, bbf) {
        MouseMove 372, 504
        Sleep 300
        Click
        Sleep 300
        Send "!d"
        Sleep 300
        loop roomRates.Length {
            Send Format("{Text}{1}", roomRates[A_Index])
            Send "{Down}"
            Sleep 200
            ; TODO: action: set bbf if included
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
    }

    dateTimeEntry()
    Sleep 1000
    roomQtyEntry(infoObj["roomQty"])
    Sleep 1000
    commentOrderIdEntry(infoObj["orderId"], comment)
    Sleep 1000
    roomRatesEntry(infoObj["roomRates"], infoObj["bbf"])
    Sleep 1000
}


; standard RH_Macros template:
; RH_Agent(infoObj, addFromConf) {
;     roomTypeRef := Map([])
;     roomType := ""
;     for k, v in roomTypeRef {
;         if (infoObj["roomType"] = k) {
;             roomType := v
;         }
;     }
;     addFrom := "" ; add from which template pms booking(get from module ui)
;     comment := "" ; varies depends on agents, roomtypes and bbfs.

;     ; private guest names handler
;     ; should return something like [lastName, firstName, hanziName]
;     ; regex : ^[\u4e00-\u9fa5，。；：“”‘’！？、《》（）&#8203;``【oaicite:1】``&#8203;]+$


;     loop infoObj["roomQty"] {
;         Sleep 1000
;         commonEntries(infoObj, roomType, comment)
;     }
; }
