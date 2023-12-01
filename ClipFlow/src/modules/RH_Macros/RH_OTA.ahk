#Include "../../lib/DictIndex.ahk"
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

; to-be-test
profileEntry(guestName) {
    BlockInput true
    WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
    CoordMode "Mouse", "Screen"
    Sleep 2000
    MouseMove 471, 217
    Click
    Sleep 3000
    Send "!n"
    Sleep 100
    MouseMove 432, 285
    Sleep 10
    Click 3
    Sleep 10
    Send Format("{Text}{1}", guestName[1])
    MouseMove 399, 312
    Sleep 10
    Click 3
    Sleep 10
    Send Format("{Text}{1}", guestName[2])
    Sleep 10
    Send "!o"
    Sleep 10000

}

; commonEntries WIP
commonEntries(infoObj, roomTypeModded, comment) {

    roomQtyEntry(roomQty) {
        ; TODO: fill-in roomQty
    }
    ; to-be-test
    dateTimeEntry(checkin, checkout) {
        pmsCiDate := FormatTime(checkin, "MMddyyyy")
        pmsCoDate := FormatTime(checkout, "MMddyyyy")
        Sleep 200
        MouseMove 332, 356
        Sleep 1000
        Click "Down"
        MouseMove 178, 360
        Sleep 300
        Click "Up"
        MouseMove 172, 360
        Sleep 300
        Send Format("{Text}{1}", pmsCiDate)
        Sleep 100
        MouseMove 325, 398
        Sleep 300
        Click
        Sleep 300
        MouseMove 661, 543
        Sleep 300
        Click
        MouseMove 636, 543
        Sleep 300
        Click
        MouseMove 635, 543
        Sleep 300
        Click
        Sleep 300
        Click
        Sleep 300
        MouseMove 335, 405
        Sleep 300
        Click "Down"
        MouseMove 182, 409
        Sleep 300
        Click "Up"
        MouseMove 207, 415
        Sleep 300
        Send Format("{Text}{1}", pmsCoDate)
        Sleep 200
    }
    ; to-be-test
    commentOrderIdEntry(orderId, comment) {
        ; fill-in comment
        Sleep 100
        MouseMove 622, 596
        Sleep 200
        Click "Down"
        MouseMove 1140, 605
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", comment)
        Sleep 200
        ; fill-in orderId
        Sleep 200
        MouseMove 839, 555
        Sleep 100
        Click "Down"
        MouseMove 1107, 563
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", orderId)
        Sleep 200
    }
    ; WIP
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

    ; starter
    dateTimeEntry(infoObj["ciDate"], infoObj["coDate"])
    Sleep 2000
    roomQtyEntry(infoObj["roomQty"])
    Sleep 2000
    commentOrderIdEntry(infoObj["orderId"], comment)
    Sleep 2000
    roomRatesEntry(infoObj["roomRates"], infoObj["bbf"])
    Sleep 2000
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
