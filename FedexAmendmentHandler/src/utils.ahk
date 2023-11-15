getCtrlByName(vName, ctrlArray){
    loop ctrlArray.Length {
        if (vName = ctrlArray[A_Index].Name) {
            return ctrlArray[A_Index]
        }
    }
}

cleanReload(quit := 0){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    if (WinExist("旅客信息")) {
        WinSetAlwaysOnTop false, "旅客信息"
    }
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    CoordMode "Mouse", "Screen"
    if (quit = "quit") {
        ExitApp
    }
    Reload
}

quitApp(*) {
    quitConfirm := MsgBox("是否确定退出Fedex Amendment Handler? ", popupTitle, "OKCancel 4096")
    if (quitConfirm = "OK") {
        cleanReload("quit")
    } else {
        cleanReload()
    }
}

; FAH functions
getReformatDate(fedexDate){
    monthMap := Map(
        "Jan", "01",
        "Feb", "02",
        "Mar", "03",
        "Apr", "04",
        "May", "05",
        "Jun", "06",
        "Jul", "07",
        "Aug", "08",
        "Sep", "09",
        "Oct", "10",
        "Nov", "11",
        "Dec", "12",    
    )
    year := "20" . SubStr(fedexDate, 6, 2)
    month := monthMap[SubStr(fedexDate, 3, 3)]
    day := SubStr(fedexDate, 1, 2)
    return year . month . day
}

get24Hr(time, meridiem){
    if (meridiem = "AM") {
        return time
    } else {
        newTimeArray := StrSplit(time, ":")
        h := newTimeArray[1] + 12
        h := (h = 24) ? "00" : h
        m := newTimeArray[2]
    return h . ":" .  m
    }
}

getStaffNames(staffInfo){
    if (InStr(staffInfo, ", ")) {
        ; two staffs
        staffNames := []
        staffInfoSplit := StrSplit(staffInfo, ", ")
        loop staffInfoSplit.Length {
            staffInfoArr := StrSplit(staffInfoSplit[A_Index], " ")
            staffName := staffInfoArr[3] . " " . StrSplit(staffInfoArr[4], "(")[1]
            staffNames.Push(staffName)
        }
    } else {
        ; one staff
        staffInfoArr := StrSplit(staffInfo, " ")
        staffName := , staffInfoArr[3] . " " . StrSplit(staffInfoArr[4], "(")[1]
        staffNames.Push(staffName)
    }
    return staffNames
}

getStayHours(inboundDate, inboundTime, outboundDate, outboundTime){
    ib := inboundDate . StrSplit(inboundTime, ":")[1] . StrSplit(inboundTime, ":")[2]
    ob := outboundDate . StrSplit(outboundTime, ":")[1] . StrSplit(outboundTime, ":")[2]
    stayMinutes := DateDiff(ob, ib, "M")
    ; return stayMinutes
    hour := Floor(stayMinutes / 60)
    minute := Integer((stayMinutes / 60 - hour) * 60)
    h := hour < 10 ? "0" . hour : hour
    m := minute < 10 ? "0" . minute : minute
    return h . ":" . m
}

getDaysActual(hoursAtHotel) {
    h := StrSplit(hoursAtHotel, ":")[1]
    m := StrSplit(hoursAtHotel, ":")[2]
    if (h < 24) {
        return 1
    } else if (Mod(h, 24) = 0 && m = 0) {
        return Integer(h / 24)
    } else if (h >= 24 || m != 0) {
        return Integer(h / 24 + 1)
    }
}