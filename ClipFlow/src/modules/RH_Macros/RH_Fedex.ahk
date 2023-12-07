#Include "../../lib/utils.ahk"
#Include "../../lib/DictIndex.ahk"
; FedEx WIP
;KNOWN ISSUE: when modifying daily detail, error popup needs to handle
; WIP: CHANGE
class FedexEntry {
    static USE(infoObj, initX:=194, initY:=183) {
        ; CoordMode "Mouse", "Screen"
        pmsCiDate := (StrSplit(infoObj["ETA"], ":"))[1] < 10
            ? DateAdd(infoObj["ciDate"], -1, "days")
            : infoObj["ciDate"]
        pmsCoDate := infoObj["coDate"]
        pmsNts := DateDiff(pmsCoDate, pmsCiDate, "days")

        this.profileEntry(infoObj["crewNames"])
        Sleep 500
        if (infoObj["resvType"] = "ADD") {
            this.roomQtyEntry(infoObj["crewNames"].Length)
        }
        Sleep 500
        this.dateTimeEntry(infoObj["ciDate"], infoObj["coDate"], infoObj["ETA"], infoObj["ETD"])
        Sleep 500
        this.moreFieldsEntry(infoObj)
        Sleep 500
        this.commentTripNumTrackingEntry(infoObj)
        Sleep 500
        pmsNts := DateDiff(pmsCoDate, pmsCiDate, "days")
        if (infoObj["daysActual"] < pmsNts) {
            this.dailyDetailsEntry(infoObj["daysActual"])
        }
        Sleep 500
        FedexEntry.rateCodeEntry()
    }
    ; tested
    static profileEntry(crewNames, initX := 471, initY := 217) {
        if (A_Index = 0) {
            A_Index := 1
        }
        crewName := StrSplit(crewNames[A_Index], " ")
        Sleep 1000
        MouseMove initX, initY ; 471, 217
        Click
        Sleep 3000
        Send "!n"
        Sleep 100
        Send "{Esc}"
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
    ; tested
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
            Sleep 100
        }
        Sleep 100
    }
    ; tested. seems that rate code entry is not neccessary
    static dateTimeCodeEntry(checkin, checkout, ETA, ETD, initX := 323, initY := 506) {
        pmsCiDate := (StrSplit(ETA, ":")[1]) < 10
            ? FormatTime(DateAdd(checkin, -1, "days"), "MMddyyyy")
            : FormatTime(checkin, "MMddyyyy")
        pmsCoDate := FormatTime(checkout, "MMddyyyy")
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
        Send "{Enter}"
        Sleep 100
        loop 3 {
            Send "{Esc}"
            Sleep 100
        }
        ; fill in ETA & ETD
        MouseMove initX - 29, initY + 91 ; 294, 597
        Sleep 100
        ; Click
        ; Sleep 100
        ; Send "{Enter}"
        ; Sleep 100
        ; Send "{Enter}"
        ; Sleep 100
        ; Send "{Enter}"
        ; Sleep 100
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
    ; tested
    static commentTripNumTrackingEntry(infoObj, initX := 622, initY := 589) {
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
            Send "^x"
            Sleep 100
            prevComment := A_Clipboard
            comment := Format("Changed to {1}={2}day(s), New Stay:{3}-{4} // Before Update:{5}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"], prevComment)
        }
        ; fill-in comment
        Sleep 100
        MouseMove initX, initY ; 622, 596
        Sleep 100
        ; Click "Down"
        ; MouseMove initX + 518, initY + 36 ; 1140, 605
        ; Sleep 100
        ; Click "Up"
        Click
        Sleep 100
        Send Format("{Text}{1}", comment)
        Sleep 100
        ; fill-in new flight and trip
        MouseMove initX + 307, initY - 35 ; 929, 554
        Sleep 100
        Click 3
        Sleep 100
        Send Format("{Text}{1}  {2}", infoObj["flightIn"], infoObj["tripNum"])
        Sleep 100
        ; fill-in tracking
        MouseMove initX + 301, initY -84 ; 923, 505
        Sleep 100
        Click 3
        Sleep 100
        ; entry tracking number at field "contact", easier to access
        Send Format("{Text}Tracking:{1}", infoObj["tracking"])
        Sleep 100
    }
    ; tested
    static moreFieldsEntry(infoObj, initX := 236, initY := 333) {
        schdCiDate := FormatTime(infoObj["ciDate"], "MMddyyyy")
        schdCoDate := FormatTime(infoObj["coDate"], "MMddyyyy")
        MouseMove initX, initY ; 236, 333
        Sleep 100
        Click
        Sleep 100
        MouseMove initX + 450, initY + 126 ; 686, 459
        Sleep 200
        Click "Down"
        MouseMove initX + 242, initY + 126 ; 478, 459
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
    ; tested(change) // to-be-test(add)
    static dailyDetailsEntry(daysActual, initX := 372, initY := 524) {
        MouseMove initX, initY ; 372, 524
        Sleep 100
        Click
        Sleep 100
        Send "!d"
        Sleep 1500
        loop daysActual {
            Send "{Down}"
            Sleep 200
        }
        Send "!e"
        Sleep 100
        loop 4 {
            Send "{Tab}"
            Sleep 200
        }
        Send "{Text}NRR"
        Sleep 100
        Send "!o"
        loop 3 {
            Send "{Enter}"
            Sleep 200
        }
        Sleep 300
        Send "!o"
        Sleep 100
        loop 5 {
            Send "{Escape}"
            Sleep 200
        }
        Sleep 100
    }
    ; WIP
    static splitParty(crewNames, initX:=456, initY:=482) {
        Send "!t"
        Sleep 100
        MouseMove initX, initY
        Sleep 100
        Send Click
        Sleep 100
        Send "!s" ; !s: Split; !a: Split All
        Sleep 100
        Send "!r"
        Sleep 1000
    }
    ; to-be-test 
    static rateCodeEntry(initX := 326, initY := 507){
        MouseMove initX, initY
        Sleep 100
        Click "Down"
        Sleep 100
        MouseMove initX - 71, initY
        Sleep 100
        Click "Up"
        Sleep 100
        Send "{Text}FEDEXN"
        Sleep 100
        Send "{Tab}"
        Sleep 100
        loop 3 {
            Send "{Esc}"
            Sleep 100
        }
    }
}

RH_Fedex(infoObj) {
    ; infoObj["ETA"] := "12:52"
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
            notifierOC := Format("
                (
                    待修改订单数量：{1}, 已完成{2}个。
                    请先打开需要修改的 Fedex 预订。

                    确定(Enter)：     开始修改预订
                    取消(Esc)：       退出修改
                )", roomQty, A_Index - 1)
            notifierYNC :=  Format("
                (
                    待修改订单数量：{1}, 已完成{2}个。
                    请先打开需要修改的 Fedex 预订。

                    是(Yes)：         开始修改预订
                    否(No):           跳过第一个修改    
                    取消(Cancel):     退出修改
                )", roomQty, A_Index - 1)

            options := roomQty = 2 ? "YesNoCancel" : "OKCancel"
            notifier := roomQty = 2 ? notifierYNC : notifierOC
            if (A_Index = 2) {
                notifier := notifierOC
                options := "OKCancel"
            }
            changePopup := MsgBox(notifier, "RH-FedEx - CHANGE", options . " 4096")
            if (changePopup = "Cancel") {
                Utils.cleanReload(winGroup)
            } else if (changePopup = "No") {
                continue
            } else if (changePopup = "Yes"){
                FedexEntry.USE(infoObj)
            } else {
                FedexEntry.USE(infoObj)
            }
            ; leave close and save manually ???
            if (A_Index = roomQty) {
                MsgBox("已完成所有修改。","RH-FedEx - CHANGE", "T1 4096")
            }
        }
        Utils.cleanReload(winGroup)
    }
}