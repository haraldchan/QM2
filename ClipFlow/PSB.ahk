; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\pinyin.ahk"
#Include "../utils.ahk"
#Include "../DictIndex.ahk"

class PSB {
    static popupTitle := "ClipFlow - PSB"

    static Copy() {
        ; Coord mode needs to be set to Client or Window (guest id info window)
        global guestType := InputBox("
            (
            请输入需要复制的旅客类型：

            1 - 内地旅客
            2 - 港澳台旅客
            3 - 国外旅客
            )", this.popupTitle)
        if (guestType.Result = "Cancel") {
            cleanReload()
        }
        ; TODO: try get guest type with getpixelcolor(get the radio that is selected)
        return this.Capture(guestType.Value)
    }
    ; TODO: record capture actions
    static Capture(gType) {
        capturedInfo := []
        if (gType = 1) {
            ; from Mainland
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)

            ; capture: fullname
            capturedInfo.Push(A_Clipboard)
            ; capture: address
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            return this.parseGuestInfo(1, capturedInfo)
        } else if (gType = 2) {
            ; from HK/MO/TW
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)

            ; capture: fullname
            capturedInfo.Push(A_Clipboard)
            ; capture: lastname
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            return this.parseGuestInfo(2, capturedInfo)
        } else {
            ; from abroad
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)

            ; capture: lastname
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            capturedInfo.Push(A_Clipboard)
            ; capture: country
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            return guestProfile := this.parseGuestInfo(3, capturedInfo)
        }
    }

    static parseGuestInfo(gType, infoArr) {
        guestProfile := Map()
        guestProfile["birthday"] := FormatTime(infoArr[1], "MMddyyyy")
        guestProfile["gender"] := (infoArr[2] = "男") ? "Mr" : "Ms"
        guestProfile["idNum"] := infoArr[3]
        if (gType = 1) {
            ; from Mainland
            guestProfile["language"] := "C"
            guestProfile["country"] := "CN"
            guestProfile["altName"] := infoArr[4]
            guestProfile["lastName"] := getFullnamePinyin(infoArr[4])[1]
            guestProfile["firstName"] := getFullnamePinyin(infoArr[4])[2]
            if (StrLen(infoArr[4]) = 18) {
                guestProfile["idType"] := "IDC"
            } else if (StrLen(infoArr[4]) = 9) {
                guestProfile["idType"] := (SubStr(guestProfile["idNum"], 1, 1) = "C") ? "MRP" : "IDP"
            } else {
                guestProfile["idType"] := ""
            }
            guestProfile["address"] := infoArr[5]
            guestProfile["province"] := getProvince(guestProfile["address"])
        } else if (gType := 2) {
            guestProfile["language"] := "E"
            guestProfile["country"] := "CN"
            guestProfile["altName"] := infoArr[4]
            guestProfile["lastName"] := infoArr[5]
            guestProfile["firstName"] := infoArr[6]
            if (SubStr(guestProfile["idNum"], 1, 1) = "H" || SubStr(guestProfile["idNum"], 1, 1) = "M") {
                guestProfile["idType"] := "HKC"
            } else {
                guestProfile["idType"] := "TWT"
            }
        } else {
            ; from abroad
            guestProfile["language"] := "E"
            guestProfile["idType"] := "NOP"
            guestProfile["lastName"] := infoArr[4]
            guestProfile["firstName"] := infoArr[5]
            guestProfile["country"] :=  getCountryCode(infoArr[6])
        }
        return guestProfile
    }


    static Paste(guestProfileMap) {
        CoordMode "Mouse", "Screen"
        ; TODO:(advance feature) try set custom relative coords 
        ;       1.get profile window image
        ;       2.use ImageSearch to get initial(anchor) pixel pos
        ;       3.if it works, try implement it to other scripts

        ; TODO: mouse action to send(not paste) guest profile info
        ; action: common info: lastName, firstName, language, gender, country, birthday, idType, idNum
        if (guestProfileMap.Has("altName")) {
            ; action: with hanzi name
            ; province, altName, gender(in altName window)
        }
    }
}