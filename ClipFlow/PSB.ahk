; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\pinyin.ahk"
#Include "../utils.ahk"
#Include "../DictIndex.ahk"

class PSB {
    static popupTitle := "ClipFlow - PSB"
    static guestProfile := Map()
    static profileFormat := [
        "language",
        "lastName",
        "firstName",
        "altName",
        "gender",
        "address",
        "country",
        "province",
        "birthday",
        "idType",
        "idNum"
    ]

    static Capture(gType) {
        capturedInfo := []
        if (gType = 1) {
            ; from Mainland
            ; capture: fullname
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)
            ; capture: address
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            this.guestProfile["language"] := "C"
            this.guestProfile["country"] := "CN"
            this.guestProfile["altName"] := capturedInfo[1]
            this.guestProfile["lastName"] := getFullnamePinyin(capturedInfo[1])[1]
            this.guestProfile["firstName"] := getFullnamePinyin(capturedInfo[1])[2]
            this.guestProfile["gender"] := (capturedInfo[2] = "男") ? "Mr" : "Ms"
            this.guestProfile["birthday"] := FormatTime(capturedInfo[3], "MMddyyyy")
            this.guestProfile["idNum"] := capturedInfo[4]
            if (StrLen(capturedInfo[4]) = 18) {
                this.guestProfile["idType"] := "IDC"
            } else if (StrLen(capturedInfo[4]) = 9) {
                this.guestProfile["idType"] := (SubStr(this.guestProfile["idNum"], 1, 1) = "C") ? "MRP" : "IDP"
            } else {
                this.guestProfile["idType"] := ""
            }
            this.guestProfile["address"] := capturedInfo[5]
            this.guestProfile["province"] := getProvince(this.guestProfile["address"])
        } else if (gType = 2) {
            ; from HK/MO/TW
            ; capture: fullname
            capturedInfo.Push(A_Clipboard)
            ; capture: lastname
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            this.guestProfile["language"] := "E"
            this.guestProfile["country"] := "CN"
            this.guestProfile["altName"] := capturedInfo[1]
            this.guestProfile["lastName"] := capturedInfo[2]
            this.guestProfile["firstName"] := capturedInfo[3]
            this.guestProfile["gender"] := (capturedInfo[4] = "男") ? "Mr" : "Ms"
            this.guestProfile["birthday"] := FormatTime(capturedInfo[5], "MMddyyyy")
            this.guestProfile["idNum"] := capturedInfo[6]
            if (SubStr(this.guestProfile["idNum"], 1, 1) = "H" || SubStr(this.guestProfile["idNum"], 1, 1) = "M") {
                this.guestProfile["idType"] := "HKC"
            } else {
                this.guestProfile["idType"] := "TWT"
            }
        } else {
            ; from abroad
            ; capture: lastname
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            capturedInfo.Push(A_Clipboard)
            ; capture: gender
            capturedInfo.Push(A_Clipboard)
            ; capture: birthday
            capturedInfo.Push(A_Clipboard)
            ; capture: id
            capturedInfo.Push(A_Clipboard)
            ; capture: country
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            this.guestProfile["language"] := "E"
            this.guestProfile["idType"] := "NOP"
            this.guestProfile["lastName"] := capturedInfo[1]
            this.guestProfile["firstName"] := capturedInfo[2]
            this.guestProfile["gender"] := (capturedInfo[3] = "男") ? "Mr" : "Ms"
            this.guestProfile["birthday"] := FormatTime(capturedInfo[4], "MMddyyyy")
            this.guestProfile["idNum"] := capturedInfo[5]
            this.guestProfile["country"] := getCountryCode(capturedInfo[6])
        }
    }

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
        this.Capture(guestType.Value)
    }

    static Paste() {
        ; TODO: mouse action to paste guest profile info
        ; action: common info: lastName, firstName, language, gender, country, birthday, idType, idNum
        if (guestType.Value = 1) {
            ; action: with hanzi name
            ; province, altName, gender(in altName window)
        }
    }
}