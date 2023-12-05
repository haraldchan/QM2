#Include "../../lib/utils.ahk"
#Include "../../lib/DictIndex.ahk"
; FedEx WIP
;KNOWN ISSUE: when modifying daily detail, error popup needs to handle
; WIP: CHANGE
class FedexEntry {
    static USE(infoObj, initX, initY) {
        ; CoordMode "Mouse", "Screen"
        this.profileEntry(infoObj["crewNames"])
        Sleep 1000
        if (infoObj["resvType"] = "ADD") {
            this.roomQtyEntry(infoObj["crewNames"].Length)
        }
        Sleep 1000
        this.dateTimeRateCodeEntry(infoObj["ciDate"], infoObj["coDate"], infoObj["ETA"], infoObj["ETD"])
        Sleep 1000
        this.moreFieldsEntry(infoObj)
        Sleep 1000
        pmsNts := DateDiff(infoObj["coDate"], infoObj["ciDate"], "days")
        if (infoObj["daysActual"] != pmsNts) {
            this.dailyDetailsEntry(infoObj["daysActual"])
        }
    }

    static profileEntry(crewNames, initX := 471, initY := 217) {
        crewName := StrSplit(crewNames[A_Index], " ")
        Sleep 1000
        MouseMove initX, initY ; 471, 217
        Click
        Sleep 3000
        Send "!n"
        Sleep 100
        MouseMove initX - 39, initY + 68 ; 432, 285
        Sleep 10
        Click 3
        Sleep 10
        Send Format("{Text}{1}", crewName[2])
        MouseMove initX - 72, initY + 95 ; 399, 312
        Sleep 10
        Click 3
        Sleep 10
        Send Format("{Text}{1}", crewName[1])
        Sleep 10
        Send "!o"
        Sleep 2000
    }

    static roomQtyEntry(roomQty, initX := 294, initY := 441) {
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

    static dateTimeRateCodeEntry(checkin, checkout, ETA, ETD, initX := 323, initY := 506) {
        pmsCiDate := (StrSplit(ETA, ":")[1]) < 10
            ? FormatTime(DateAdd(checkin, -1, "days"), "MMddyyyy")
            : FormatTime(checkin, "MMddyyyy")
        pmsCoDate := FormatTime(checkout, "MMddyyyy")
        ; fill-in rate code
        MouseMove initX, initY ; 323, 506
        Sleep 100
        Click
        Sleep 100
        Send "{Text}FEDEXN"
        Sleep 100
        Send "{Tab}"
        Sleep 100
        MouseMove initX + 9, initY - 150 ; 332, 356
        Sleep 100
        ; fill-in checkin/checkout
        Click "Down"
        MouseMove initX - 145, initY - 146 ; 178, 360
        Sleep 100
        Click "Up"
        MouseMove initX - 151, initY - 146 ; 172, 360
        Sleep 100
        Send Format("{Text}{1}", pmsCiDate)
        Sleep 100
        MouseMove initX + 2, initY - 108 ; 325, 398
        Sleep 100
        Click
        Sleep 100
        MouseMove initX + 338, initY + 37 ; 661, 543
        Sleep 100
        Click
        MouseMove initX + 313, initY + 37 ; 636, 543
        Sleep 100
        Click
        MouseMove initX + 312, initY + 37 ; 635, 543
        Sleep 100
        Click
        Sleep 100
        Click
        Sleep 100
        MouseMove initX + 12, initY - 101 ; 335, 405
        Sleep 100
        Click "Down"
        MouseMove initX - 141, initY - 97 ; 182, 409
        Sleep 100
        Click "Up"
        MouseMove initX - 116, initY - 91 ; 207, 415
        Sleep 100
        Send Format("{Text}{1}", pmsCoDate)
        Sleep 100
        ; fill in ETA & ETD
        MouseMove initX - 29, initY + 91 ; 294, 597
        Sleep 100
        Click
        Sleep 100
        Send "{Enter}"
        Sleep 100
        Send "{Enter}"
        Sleep 100
        Send "{Enter}"
        Sleep 100
        MouseMove initX - 3, initY + 91 ; 320, 597
        Sleep 100
        Click "Down"
        MouseMove initX - 123, initY + 91 ; 200, 597
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", ETA)
        Sleep 100
        MouseMove initX + 176, initY + 91 ; 499, 597
        Sleep 100
        Click "Down"
        MouseMove initX + 7, initY + 88 ; 330, 594
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", ETD)
        Sleep 100
    }

    static commentTripNumEntry(infoObj, initX := 622, initY := 569) {
        ; set comment
        comment := ""
        if (infoObj["resvType"] = "ADD") {
            comment := Format("RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}day(s), ActualStay: {3}-{4}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"])
        } else {
            ; copy current comment and reformat
            MouseMove initX, initY ; 622, 596
            Sleep 100
            Click "Down"
            MouseMove initX + 518, initY + 36 ; 1140, 605
            Sleep 100
            Click "Up"
            Sleep 100
            Send "^c"
            prevComment := A_Clipboard
            comment := Format("Changed to {1}={2}day(s), New Stay:{3}-{4} `n Before Update:{5}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"], prevComment)
        }
        ; fill-in comment
        Sleep 100
        MouseMove initX, initY ; 622, 596
        Sleep 100
        Click "Down"
        MouseMove initX + 518, initY + 36 ; 1140, 605
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", comment)
        Sleep 100
        ; fill-in new flight and trip
        MouseMove initX + 217, initY - 50 ; 839, 555
        Sleep 100
        Click "Down"
        MouseMove initX + 485, initY - 42 ; 1107, 563
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}  {2}", infoObj["flightIn"], infoObj["tripNum"])
        Sleep 100
        ; fill-in tracking
        MouseMove initX + 281, initY - 99 ; 930, 506
        Sleep 100
        Click 3
        Sleep 100
        ; entry tracking number at field "contact", easier to access
        Send Format("{Text}{1}", infoObj["tracking"])
        Sleep 100
    }

    static moreFieldsEntry(infoObj, checkin, checkout, initX := 236, initY := 333) {
        schdCiDate := FormatTime(checkin, "MMddyyyy")
        schdCoDate := FormatTime(checkout, "MMddyyyy")
        MouseMove initX, initY ; 236, 333
        Sleep 100
        Click
        Sleep 100
        MouseMove initX := 450, initY := 126 ; 686, 459
        Sleep 200
        Click "Down"
        MouseMove initX := 242, initY + 126 ; 478, 459
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", infoObj["flightIn"])
        Sleep 100
        MouseMove initX + 436, initY + 170 ; 672, 503
        Sleep 281
        Click "Down"
        MouseMove initX + 287, initY + 170 ; 523, 503
        Sleep 358
        Click "Up"
        Send Format("{Text}{1}", schdCiDate)
        Sleep 100
        MouseMove initX + 449, initY + 193 ; 685, 526
        Sleep 100
        Click "Down"
        MouseMove initX + 186, initY + 188 ; 422, 521
        Sleep 100
        Click "Up"
        Send Format("{Text}{1}", infoObj["ETA"])
        Sleep 100
        MouseMove initX + 686, initY + 128 ; 922, 461
        Sleep 100
        Click "Down"
        MouseMove initX + 468, initY + 126 ; 704, 459
        Sleep 100
        Click "Up"
        Send Format("{Text}{1}", infoObj["flightOut"])
        Sleep 100
        MouseMove initX + 681, initY + 170 ; 917, 503
        Sleep 100
        Click "Down"
        MouseMove initX + 401, initY + 174 ; 637, 507
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", schdCoDate)
        Sleep 100
        MouseMove initX + 686, initY + 191 ; 922, 524
        Sleep 100
        Click "Down"
        MouseMove initX + 404, initY + 190 ; 640, 523
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", infoObj["ETD"])
        Sleep 100
        MouseMove initX + 605, initY + 347 ; 841, 680
        Sleep 100
        Click
        Sleep 100
    }
    ; WIP!!!!
    static dailyDetailsEntry(daysActual, initX := 372, initY := 524) {
        MouseMove initX, initY ; 372, 524
        Sleep 100
        Click
        Sleep 100
        Send "!d"
        Sleep 3000
        loop daysActual {
            Send "{Down}"
            Sleep 200
        }
        Send "0"
        Sleep 100
        MouseMove initX + 246, initY - 19 ; 618, 505
        Sleep 100
        Send "!e"
        Sleep 100
        loop 4 {
            Send "{Tab}"
            Sleep 200
        }
        Send "{Text}NRR"
        Sleep 100
        MouseMove initX + 46, initY - 127 ; 418, 397
        Sleep 100
        Send "!o"
        Sleep 100
        loop 5 {
            Send "{Escape}"
            Sleep 200
        }
        Send "!o"
        Sleep 1500
        Send "{Space}"
        Sleep 500
        MouseMove initX + 356, initY + 44 ; 728, 568
        Sleep 100
        Send "!o"
        Sleep 100
        MouseMove initX + 170, initY - 51 ; 542, 473
        Sleep 100
        Click
        MouseMove initX + 272, initY + 19 ; 644, 543
        Sleep 100
        Click
        Sleep 100
    }
    ; WIP
    static splitPartyEntry(crewNames, initX, initY) {

    }
}

RH_Fedex(infoObj) {
    roomQty := infoObj["crewNames"].Length

    ; starter
    if (infoObj["resvType"] = "CHANGE") {
        changeModification(infoObj)
    } else {
        addModification(infoObj)
    }

    ; WIP
    addModification(infoObj) {
        addFrom := InputBox("请指定需要从哪个 FedEx Block Add-On? (请输入 BlockCode )", "Reservation Handler", , SubStr(A_YYYY, 3, 4) . A_MM . A_MDay . "FEDEX")
        if (addFrom.Result = "Cancel") {
            Utils.cleanReload(winGroup)
        }

        ; TODO: action: adds on from a block pm(by using addFrom.Value), then open it.
        Sleep 1000
        Send "!r"
        Sleep 100
        Send "u"
        Sleep 3000
        Loop 8 {
            Send "{Tab}"
            Sleep 10
        }
        Send addFrom.Value
        Sleep 100
        Loop 4 {
            Send "{Tab}"
            Sleep 10
        }
        Send "{Delete}"
        Sleep 10
        Send "!h"
        Sleep 10
        Send "!t"
        Sleep 10
        Send "!o"
        Send "{Text}DKC"
        Sleep 10
        Send "!o"
        Sleep 10
        Send "{Left}"
        Sleep 10
        Send "{Enter}"
        Sleep 4000
        Send "{Esc}"
        Sleep 2500

        FedexEntry.USE(infoObj)
        Sleep 1000
        FedexEntry.splitPartyEntry(infoObj["crewNames"])

        MsgBox("已完成新增。", "Reservation Handler", "T2 4096")
    }

    ; to-be-test
    changeModification(infoObj) {
        loop roomQty {
            if (A_Index > roomQty) {
                MsgBox("已完成所有修改。")
                return
            }
            nofitier := Format("
                (
                    待修改订单数量：{1}, 已完成{2}个。
                    请先打开需要修改的 Fedex 预订。

                    确定(Enter)：     开始修改预订
                    取消(Esc)：       退出修改
                )", roomQty, A_Index)
            changePopup := MsgBox(nofitier, "Reservation Handler", "OKCancel 4096")
            if (changePopup.Result = "Cancel") {
                Utils.cleanReload(winGroup)
            }

            ; modification process
            FedexEntry.USE(infoObj)
            ; leave close and save manually ???
        }
    }
}