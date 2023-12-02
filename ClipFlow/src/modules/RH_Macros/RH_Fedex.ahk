#Include "../../lib/utils.ahk"
#Include "../../lib/DictIndex.ahk"
; FedEx WIP
;KNOWN ISSUE: when modifying daily detail, error popup needs to handle
; WIP: CHANGE
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
    ; worked
    profileEntry(infoObj) {
        crewName := StrSplit(infoObj["crewNames"][A_Index], " ")
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
        Send Format("{Text}{1}", crewName[2])
        MouseMove 399, 312
        Sleep 10
        Click 3
        Sleep 10
        Send Format("{Text}{1}", crewName[1])
        Sleep 10
        Send "!o"
        Sleep 10000
    }
    ; worked
    dateTimeRateCodeEntry(infoObj) {
        MouseMove 323, 506
        Sleep 100
        Click
        Sleep 100
        Send "{Text}FEDEXN"
        Sleep 100
        Send "{Tab}"
        Sleep 100
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
        Sleep 300
        ; }
        Sleep 2000
        ; { fill in ETA & ETD
        MouseMove 294, 597
        Sleep 200
        Click
        Sleep 200
        Send "{Enter}"
        Sleep 200
        Send "{Enter}"
        Sleep 200
        Send "{Enter}"
        Sleep 200
        MouseMove 320, 597
        Sleep 200
        Click "Down"
        MouseMove 200, 597
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", infoObj["ETA"])
        Sleep 200
        MouseMove 499, 597
        Sleep 200
        Click "Down"
        MouseMove 330, 594
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}", infoObj["ETD"])
        Sleep 200
        ; }
        Sleep 200
    }
    ; worked
    commentTripNumEntry(infoObj) {
        ; set comment
        comment := ""
        if (infoObj["resvType"] = "ADD") {
            comment := Format("RM INCL 1BBF TO CO,Hours@Hotel: {1}={2}day(s), ActualStay: {3}-{4}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"])
        } else {
            ; TODO: action: copy current comment
            MouseMove 622, 596
            Sleep 200
            Click "Down"
            MouseMove 1140, 605
            Sleep 200
            Click "Up"
            Sleep 200
            Send "^c"
            prevComment := A_Clipboard
            comment := Format("Changed to {1}={2}day(s), New Stay:{3}-{4} `n Before Update:{5}", infoObj["stayHours"], infoObj["daysActual"], infoObj["ciDate"], infoObj["coDate"], prevComment)
        }
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
        ; fill-in new flight and trip
        Sleep 200
        MouseMove 839, 555
        Sleep 100
        Click "Down"
        MouseMove 1107, 563
        Sleep 200
        Click "Up"
        Sleep 200
        Send Format("{Text}{1}  {2}", infoObj["flightIn"], infoObj["tripNum"])
        Sleep 200
        ; fill-in tracking
        MouseMove 930, 506
        Sleep 200
        Click 3
        Sleep 200
        Send Format("{Text}{1}", infoObj["tracking"])
    }
    ; worked
    moreFieldsEntry(infoObj) {
        MouseMove 236, 333
        Sleep 100
        Click
        Sleep 100
        MouseMove 686, 459
        Sleep 200
        Click "Down"
        MouseMove 478, 459
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", infoObj["flightIn"])
        Sleep 100
        MouseMove 672, 503
        Sleep 281
        Click "Down"
        MouseMove 523, 503
        Sleep 358
        Click "Up"
        Send Format("{Text}{1}", schdCiDate)
        Sleep 100
        MouseMove 685, 526
        Sleep 100
        Click "Down"
        MouseMove 422, 521
        Sleep 100
        Click "Up"
        Send Format("{Text}{1}", infoObj["ETA"])
        Sleep 100
        MouseMove 922, 461
        Sleep 100
        Click "Down"
        MouseMove 704, 459
        Sleep 100
        Click "Up"
        Send Format("{Text}{1}", infoObj["flightOut"])
        Sleep 100
        MouseMove 917, 503
        Sleep 100
        Click "Down"
        MouseMove 637, 507
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", schdCoDate)
        Sleep 100
        MouseMove 922, 524
        Sleep 100
        Click "Down"
        MouseMove 640, 523
        Sleep 100
        Click "Up"
        Sleep 100
        Send Format("{Text}{1}", infoObj["ETD"])
        Sleep 100
        MouseMove 841, 680
        Sleep 100
        Click
        Sleep 2000
    }
    ; WIP
    dailyDetailsEntry(infoObj) {
        MouseMove 372, 524
        Sleep 300
        Click
        Sleep 1000
        Send "!d"
        Sleep 3000
        loop infoObj["daysActual"] {
            Send "{Down}"
            Sleep 200
        }
        Send "0"
        Sleep 200
        MouseMove 618, 505
        Sleep 200
        Send "!e"
        Sleep 200
        loop 4 {
            Send "{Tab}"
            Sleep 500
        }
        Send "{Text}NRR"
        Sleep 500
        MouseMove 418, 397
        Sleep 500
        Send "!o"
        Sleep 500
        loop 3 {
            Send "{Escape}"
            Sleep 500
        }
        Send "!o"
        Sleep 1500
        Send "{Space}"
        Sleep 500

        MouseMove 728, 568
        Sleep 300
        Send "!o"
        Sleep 300
        MouseMove 542, 473
        Sleep 300
        Click
        MouseMove 644, 543
        Sleep 300
        Click
        Sleep 300
        ; }
        Sleep 2000
    }
    ; worked
    addModification(infoObj) {
        addFrom := InputBox("请指定需要从哪个 Block Add-On? (请输入 BlockCode )", "Reservation Handler", , SubStr(A_YYYY, 3, 4) . A_MM . A_MDay . "FEDEX")
        if (addFrom.Result = "Cancel") {
            Utils.cleanReload(winGroup)
        }

        loop infoObj["roomQty"] {
            if (A_Index > infoObj["roomQty"]) {
                MsgBox("已完成所有新增。")
                return
            } else {
                ; TODO: action: adds on from a block pm(by using addFrom.Value), then open it.
                Sleep 1000
                Send "!r"
                Sleep 100
                Send "u"
                Sleep 3000
                Loop 8 {
                    Send "{Tab}"
                    Sleep 100
                }
                Send addFrom.Value
                Sleep 100
                Loop 4 {
                    Send "{Tab}"
                    Sleep 100
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

                profileEntry(infoObj)
                Sleep 1000
                dateTimeRateCodeEntry(infoObj)
                Sleep 1000
                commentTripNumEntry(infoObj)
                Sleep 1000
                moreFieldsEntry(infoObj)
                Sleep 1000
                if (infoObj["daysActual"] != pmsNts) {
                    dailyDetailsEntry(infoObj)
                    Sleep 1000
                }
                Send "!o"
                Sleep 5000
                Send "{Enter}"
                ; TODO: action: close and save the reservation.
                Sleep 5000
                Loop 5 {
                    Send "!c"
                    Sleep 100
                }
                Loop 2 {
                    Send "{Esc}"
                    Sleep 100
                }
            }
        }
    }
    ; WIP
    changeModification(infoObj) {
        loop infoObj["roomQty"] {
            if (A_Index = infoObj["roomQty"]) {
                MsgBox("已完成所有修改。")
                return
            }
            nofitier := Format("待修改订单数量：{1}, 已完成{2}个。`n`n 请先打开需要修改的 Fedex 预订。", infoObj["roomQty"], A_Index)
            changeFrom := MsgBox(nofitier, "Reservation Handler", "OKCancel 4096")
            if (changeFrom.Result = "Cancel") {
                Utils.cleanReload(winGroup)
            }
            profileEntry(infoObj)
            Sleep 1000
            dateTimeRateCodeEntry(infoObj)
            Sleep 1000
            commentTripNumEntry(infoObj)
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
