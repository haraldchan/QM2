; #Include "../../lib/DictIndex.ahk"
#Include "../../../../Lib/ClipFlow/DictIndex.ahk"

class Entry {
    ; the initX, initY for USE() should be top-left corner of current booking window
    static USE(curTemplate, infoObj, roomType, comment, pmsGuestNames, initX := 193, initY := 182) {
        MsgBox("Start in seconds...", "Reservation Handler", "T1 4096")
        Entry.addFromTemplates(curTemplate)
        ; Sleep 1000
        Entry.profileEntry(pmsGuestNames[1])
        ; sleep 1000
        Entry.roomQtyEntry(infoObj["roomQty"])
        ; sleep 1000
        Entry.roomTypeEntry(roomType)
        ; sleep 1000
        Entry.dateTimeEntry(infoObj["ciDate"], infoObj["coDate"])
        ; sleep 1000
        Entry.commentOrderIdEntry(infoObj["orderId"], comment)
        ; sleep 1000
        if (!(Utils.arrayEvery((item) => item = 0, infoObj["bbf"]))) {
            Entry.breakfastEntry(infoObj["bbf"])
        }
        ; sleep 1000
        Entry.roomRatesEntry(infoObj["roomRates"])
        sleep 1000
        MsgBox("Completed.", "Reservation Handler", "T2 4096")
        ; Utils.cleanReload(winGroup)
    }

    ; tested
    static addFromTemplates(agentTemplate) {
        Send "!r"
        sleep 100
        Send "u"
        Sleep 500
        Send Format("{Text}{1}", agentTemplate)
        Sleep 100
        Send "!h"
        Sleep 100
        Send "!t"
        Sleep 100
        Send "!o"
        Sleep 100
        Send "!o"
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
    }

    ; tested
    static profileEntry(guestName, initX := 471, initY := 217) {
        trayTip "录入中：Profile"
        Sleep 1000
        MouseMove initX, initY ;471, 217
        Click
        Sleep 3000
        Send "!n"
        Sleep 100
        MouseMove initX - 39, initY + 68 ;432, 285
        Sleep 100
        Click 3
        Sleep 100
        Send Format("{Text}{1}", guestName[1])
        Sleep 100
        MouseMove initX - 72, initY + 95 ;399, 312
        Sleep 100
        Click 3
        Sleep 100
        Send Format("{Text}{1}", guestName[2])
        Sleep 100
        if (guestName.Length = 3) {
            MouseMove initX - 17, initY + 67 ; 454, 284
            Sleep 100
            Click 3
            Sleep 3000
            Send Format("{Text}{1}", guestName[3])
            Sleep 100
            Send "!o"
        }
        Sleep 1500
        Send "!o"
        Sleep 2000
    }

    ; tested
    static roomQtyEntry(roomQty, initX := 294, initY := 441) {
        trayTip "录入中：房间数量"
        ; TODO: fill-in roomQty
        MouseMove initX, initY
        Sleep 100
        Click 3
        Sleep 100
        Send Format("{Text}{1}", roomQty)
        Sleep 100
        Send "{Tab}"
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
    }

    ; tested
    static roomTypeEntry(roomType, initX := 472, initY := 465) {
        trayTip "录入中：房型"
        MouseMove initX, initY
        Sleep 100
        Click 2
        Sleep 100
        Send roomType
        Sleep 100
        Send "{Tab}"
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
        MouseMove initX - 152, initY  ;320, 461
        Sleep 100
        Click "Down"
        Sleep 100
        MouseMove initX - 232, initY ; 240, 465
        Sleep 100
        Click "Up"
        Sleep 100
        Send roomType
        Sleep 100
        Send "{Tab}"
        Sleep 100
    }

    ; tested
    static dateTimeEntry(checkin, checkout, initX := 332, initY := 356) {
        trayTip "录入中：入住/退房日期"
        pmsCiDate := FormatTime(checkin, "MMddyyyy")
        pmsCoDate := FormatTime(checkout, "MMddyyyy")
        Sleep 100
        MouseMove initX, initY ;332, 356
        Sleep 1000
        Click "Down"
        MouseMove initX - 154, initY + 4 ;178, 360
        Sleep 100
        Click "Up"
        MouseMove initX - 160, initY + 4 ;172, 360
        Sleep 100
        Send Format("{Text}{1}", pmsCiDate)
        Sleep 100
        MouseMove initX - 7, initY + 42 ;325, 398
        Sleep 100
        Click
        Sleep 100
        MouseMove initX + 329, initY + 187 ;661, 543
        Sleep 100
        Click
        MouseMove initX + 304, initY + 187 ;636, 543
        Sleep 100
        Click
        MouseMove initX + 303, initY + 187 ;635, 543
        Sleep 100
        Click
        Sleep 100
        Click
        Sleep 100
        MouseMove initX + 3, initY + 49 ;335, 405
        Sleep 100
        Click "Down"
        MouseMove initX - 150, initY + 53 ;182, 409
        Sleep 100
        Click "Up"
        MouseMove initX - 125, initY + 59 ;207, 415
        Sleep 100
        Send Format("{Text}{1}", pmsCoDate)
        Sleep 100
        Send "{Tab}"
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
    }

    ; tested
    static commentOrderIdEntry(orderId, comment, initX := 622, initY := 596) {
        trayTip "录入中：COMMENT，订单号"
        ; fill-in comment
        Sleep 100
        MouseMove initX, initY ;622, 596
        Sleep 100
        Click "Down"
        MouseMove initX + 518, initY + 9 ;1140, 605
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", comment)
        Sleep 100
        ; fill-in orderId
        Sleep 100
        MouseMove initX + 217, initY - 41 ;839, 555
        Sleep 100
        Click "Down"
        MouseMove initX + 485, initY - 33 ;1107, 563
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", orderId)
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
    }

    ; tested?
    static roomRatesEntry(roomRates, initX := 372, initY := 524) {
        trayTip "录入中：房价"
        MouseMove initX, initY ;372, 504
        Sleep 100
        Click
        Sleep 100
        loop 5 {
            Send "{Esc}"
            Sleep 200
        }
        Sleep 100
        Send "!d"
        Sleep 100
        loop roomRates.Length {
            Send Format("{Text}{1}", roomRates[A_Index])
            Send "{Down}"
            Sleep 100
            ; TODO: action: set bbf if included
        }
        MouseMove initX + 356, initY + 44 ;728, 548
        Sleep 100
        Send "!o"
        Sleep 500
        Send "{Esc}"
        Sleep 100
    }
    ; tested
    static breakfastEntry(bbf, initX := 352, initY := 548) {
        trayTip "录入中：早餐"
        ;entry bbf package
        Sleep 100
        MouseMove initX, initY
        Sleep 100
        Click
        Sleep 2000
        Send "!n"
        Sleep 100
        Send "{Text}BFNP"
        Sleep 100
        Send "!o"
        Sleep 1000
        Send "!o"
        Sleep 100
        Send "{Esc}"
        Sleep 500
        ; change "Adults"
        MouseMove initX - 67, initY - 124 ; 285, 424
        Sleep 100
        Click 3
        Send Format("{Text}{1}", bbf[1])
        ; Send "1"
        Sleep 100
    }
    ; WIP
    static saveBooking(initX, initY) {
        ;TODO: action: save modified booking, handle popups.
    }

    ; WIP
    static splitPartyEntry(guestNames, initX, initY) {
        ;TODO: action: split party
    }


}

; Kingsley WIP
RH_Kingsley(infoObj, addFromConf) {
    ; template booking name
    thisTemplate := "template-kingsley"
    ; convert roomType
    roomTypeRef := Map(
        "标准大床房", "SKC",
        "标准双床房", "STC",
        "豪华城景大床房", "DKC",
        "豪华城景双床房", "DTC",
        "豪华江景大床房", "DKR",
        "豪华江景双床房", "DTR",
        "行政豪华城景大床房", "CKC",
        "行政豪华城景双床房", "CTC",
        "行政豪华江景大床房", "CKR",
        "行政豪华江景双床房", "CTR"
        "行政尊贵套房", "CSK"
    )
    roomType := ""
    for k, v in roomTypeRef {
        if (infoObj["roomType"] = k) {
            roomType := v
        }
    }

    ; define breakfast comment
    breakfastType := (SubStr(roomType, 1, 1) = "C") ? "CBF" : "BBF"
    breakfastQty := infoObj["bbf"][1]
    comment := (breakfastQty = 0)
        ? "RM TO TA"
        : Format("RM INCL {1}{2} TO TA", breakfastQty, breakfastType)

    ; reformat guest names
    pmsGuestNames := []
    loop infoObj["guestNames"].Length {
        curGuestName := infoObj["guestNames"][A_Index]
        if (RegExMatch(curGuestName, "^[a-zA-Z/]+$") > 0) {
            ; if only includes English alphabet, push [lastName, firstName]
            pmsGuestNames.Push(StrSplit(curGuestName, "/"))
        } else {
            pmsGuestNames.Push([
                getFullnamePinyin(curGuestName)[1], ; lastName pinyin
                getFullnamePinyin(curGuestName)[2], ; firstName pinyin
                curGuestName, ; hanzi-name
            ])
        }
    }

    ; Main booking modification
    Entry.USE(
        thisTemplate,
        infoObj,
        roomType,
        comment,
        pmsGuestNames
    )
}

; Agoda WIP
RH_Agoda(infoObj, addFromConf) {
    roomTypeRef := Map(
        "Standard Room - Queen bed", "SKC",
        "Standard Room - Twin Bed", "STC",
        "Deluxe City View with King Bed", "DKC",
        "Deluxe City View Twin Bed", "DTC",
        "Deluxe River View King Bed", "DKR"
        "Deluxe River View Twin Bed", "DTR",
        "Club Deluxe City View", "CKC",
        "Club Deluxe River View Room", "CKR",
        "Club Premier Suite", "CPK",
    )

    roomType := ""
    for k, v in roomTypeRef {
        if (infoObj["roomType"] = k) {
            roomType := v
        }
    }

    breakfastType := (roomType = "CKC" || roomType = "CKR") ? "CBF" : "BBF"
    comment := Format("ROOM INCL {1}{2} PAY BY AGODA", infoObj["bbf"][1], breakfastType)
}