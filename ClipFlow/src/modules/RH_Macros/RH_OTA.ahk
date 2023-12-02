#Include "../../lib/DictIndex.ahk"
class Entry {
    ; the initX, initY for USE() should be top-left corner of current booking window
    static USE(infoObj, roomType, comment, initX, initY) {
        this.dateTimeEntry(infoObj["ciDate"], infoObj["coDate"])
        Sleep 2000
        this.roomQtyEntry(infoObj["roomQty"])
        Sleep 2000
        this.commentOrderIdEntry(infoObj["orderId"], comment)
        Sleep 2000
        this.roomRatesEntry(infoObj["roomRates"])
        Sleep 2000
        if (!Utils.arrayEvery((item) => item = 0, infoObj["bbf"])) {
            this.breakfastEntry(infoObj["bbf"])
        }
        Sleep 2000
    }

    ; to-be-test
    static profileEntry(guestName, initX := 471, initY := 217) {
        BlockInput true
        WinSetAlwaysOnTop true, "ahk_class SunAwtFrame"
        CoordMode "Mouse", "Screen"
        Sleep 2000
        MouseMove initX, initY ;471, 217
        Click
        Sleep 3000
        Send "!n"
        Sleep 100
        MouseMove initX - 39, initY + 68 ;432, 285
        Sleep 10
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestName[1])
        Sleep 100
        MouseMove initX - 72, initY + 95 ;399, 312
        Sleep 10
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestName[2])
        Sleep 10
        Send "!o"
        Sleep 10000

    }
    ; WIP
    static splitNameEntry(guestNames, initX, initY) {
        ;TODO: action: split party

    }
    ; WIP
    static roomQtyEntry(roomQty, initX, initY) {
        ; TODO: fill-in roomQty
    }
    ; to-be-test
    static dateTimeEntry(checkin, checkout, initX := 332, initY := 356) {
        pmsCiDate := FormatTime(checkin, "MMddyyyy")
        pmsCoDate := FormatTime(checkout, "MMddyyyy")
        Sleep 200
        MouseMove initX, initY ;332, 356
        Sleep 1000
        Click "Down"
        MouseMove initX - 154, initY + 4 ;178, 360
        Sleep 300
        Click "Up"
        MouseMove initX - 160, initY + 4 ;172, 360
        Sleep 300
        Send Format("{Text}{1}", pmsCiDate)
        Sleep 100
        MouseMove initX - 7, initY + 42 ;325, 398
        Sleep 300
        Click
        Sleep 300
        MouseMove initX + 329, initY + 187 ;661, 543
        Sleep 300
        Click
        MouseMove initX + 304, initY + 187 ;636, 543
        Sleep 300
        Click
        MouseMove initX + 303, initY + 187 ;635, 543
        Sleep 300
        Click
        Sleep 300
        Click
        Sleep 300
        MouseMove initX + 3, initY + 49 ;335, 405
        Sleep 300
        Click "Down"
        MouseMove initX - 150, initY + 53 ;182, 409
        Sleep 300
        Click "Up"
        MouseMove initX - 125, initY + 59 ;207, 415
        Sleep 300
        Send Format("{Text}{1}", pmsCoDate)
        Sleep 200
    }
    ; to-be-test
    static commentOrderIdEntry(orderId, comment, initX := 622, initY := 596) {
        ; fill-in comment
        Sleep 100
        MouseMove initX, initY ;622, 596
        Sleep 200
        Click "Down"
        MouseMove initX + 518, initY + 9 ;1140, 605
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", comment)
        Sleep 200
        ; fill-in orderId
        Sleep 200
        MouseMove initX + 217, initY - 41 ;839, 555
        Sleep 100
        Click "Down"
        MouseMove initX + 485, initY - 33 ;1107, 563
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", orderId)
        Sleep 200
    }
    ; WIP
    static roomRatesEntry(roomRates, initX := 372, initY := 504) {
        MouseMove initX, initY ;372, 504
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
        MouseMove initX + 356, initY + 44 ;728, 548
        Sleep 300
        Send "!o"
        Sleep 300
        ; ⬇️ maybe won't need it
        ; MouseMove initX + 170, initY - 51 ;542, 453
        ; Sleep 300
        ; Click
        ; MouseMove initX + 272, initY + 19 ;644, 523
        ; Sleep 300
        ; Click
        Sleep 2000
    }

    static breakfastEntry(bbf, initX, initY) {
        ;TODO: action: entry bbf package
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
    Entry.profileEntry(pmsGuestNames[1])
    Sleep 1000
    Entry.USE(infoObj, roomType, comment)
    Sleep 1000
    ; TODO: close and save
    ; TODO: if roomQty > 1, split and fill-in other names
    Entry.splitNameEntry(pmsGuestNames)
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
}