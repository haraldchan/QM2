; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\pinyin.ahk"
#Include "../utils.ahk"
#Include "../pinyin.ahk"

class PSB {
    static popupTitle := "ClipFlow - PSB"
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

    static Copy() {
        global guestProfile := Map()
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
        if (guestType.Value = 1) {
            ; from China Mainland
            this.guestProfile["language"] := "C"
            this.guestProfile["Country"] := "CN"
            ; TODO: mouse action to copy fullname
            this.guestProfile["altName"] := A_Clipboard
            this.guestProfile["lastName"] := getFullnamePinyin(A_Clipboard)[1]
            this.guestProfile["lastName"] := getFullnamePinyin(A_Clipboard)[2]
            ; TODO: mouse action to copy id
            this.guestProfile["idNum"] := A_Clipboard
            if (StrLen(this.guestProfile["idNum"]) = 18) {
                this.guestProfile["idType"] := "IDC"
                this.guestProfile["birthday"] := FormatTime(SubStr(this.guestProfile["id"], 7, 8), "MMddyyyy")
                this.guestProfile["gender"] := (Mod(SubStr(this.guestProfile["id"], 17, 1), 2) = 0) ? "Ms" : "Mr"
            } else if (StrLen(this.guestProfile["idNum"]) = 9) {
                this.guestProfile["idType"] := (SubStr(this.guestProfile["idNum"], 1, 1) = "C") ? "MRP" : "IDP"
                ; TODO: mouse action to copy birthday
                this.guestProfile["birthday"] := FormatTime(A_Clipboard, "MMddyyyy")
                ; TODO: mouse action to copy gender
                this.guestProfile["gender"] := (A_Clipboard = "男") ? "Mr" : "Ms"
            } else {
                return
            }
            ; mouse action to copy address
            this.guestProfile["address"] := A_Clipboard
            this.guestProfile["province"] := getProvince(this.guestProfile["address"])
        } else if (guestType.Value = 2) {
            ; from Hongkong/Macau/Taiwan
            this.guestProfile["language"] := "E"
            this.guestProfile["country"] := "CN"
            ; TODO: mouse action to copy id
            this.guestProfile["idNum"] := A_Clipboard
            if (SubStr(this.guestProfile["idNum"], 1, 1) = "H" || SubStr(this.guestProfile["idNum"], 1, 1) = "M") {
                this.guestProfile["idType"] := "HKC"
            } else {
                this.guestProfile["idType"] := "TWT"
            }
            ; TODO: mouse action to copy last name
            this.guestProfile["lastName"] := A_Clipboard
            ; TODO: mouse action to copy first name
            this.guestProfile["firstName"] := A_Clipboard
            ; TODO: mouse action to copy full(hanzi) name
            this.guestProfile["altName"] := A_Clipboard
            ; TODO: mouse action to copy birthday
            this.guestProfile["birthday"] := FormatTime(A_Clipboard, "MMddyyyy")
            ; TODO: mouse action to copy gender
            this.guestProfile["gender"] := (A_Clipboard = "男") ? "Mr" : "Ms"
        } else {
            ; from abroad
            this.guestProfile["language"] := "E"
            this.guestProfile["idType"] := "NOP"
            ; TODO: mouse action to copy id
            this.guestProfile["idNum"] := A_Clipboard
            ; TODO: mouse action to copy last name
            this.guestProfile["lastName"] := A_Clipboard
            ; TODO: mouse action to copy first name
            this.guestProfile["firstName"] := A_Clipboard
            ; TODO: mouse action to copy country 
            this.guestProfile["country"] := getCountryCode(A_Clipboard)
            ; TODO: mouse action to copy id
            this.guestProfile["idNum"] := A_Clipboard
            ; TODO: mouse action to copy birthday
            this.guestProfile["birthday"] := FormatTime(A_Clipboard, "MMddyyyy")
            ; TODO: mouse action to copy gender
            this.guestProfile["gender"] := (A_Clipboard = "男") ? "Mr" : "Ms"
        }


    }

    static Paste() {
        ; TODO: mouse action to paste guest profile info
        if (guestType.Value = 3) {
            ; no hanzi name
        } else {
            ; has hanzi name
        }
    }
}