; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\pinyin.ahk"
#Include "../utils.ahk"
#Include "../DictIndex.ahk"

class PSB {
    static popupTitle := "ClipFlow - PSB"

    static Copy() {
        ; guestType := 0
        CoordMode "Mouse", "Window"
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
        Sleep 500
        ; TODO: try get guest type with getpixelcolor(get the radio that is selected), need more precise pixel
        return this.Capture(guestType.Value)
    }
    ; TODO: record capture actions
    static Capture(gType) {
        if (WinExist("旅客信息")){
            WinSetAlwaysOnTop true, "旅客信息"
        }
        capturedInfo := []
        ; capture: birthday
        MouseMove 755, 147
        Click "Down"
        Sleep 100
        MouseMove 755, 147
        Click "Up"
        Sleep 100
        Send "^c"
        Sleep 100
        capturedInfo.Push(A_Clipboard)
        ; capture: gender
        MouseMove 565, 147
        Sleep 100
        Click 
        Sleep 100
        Click "Right"
        Sleep 200
        Send "{c}"
        Sleep 100
        Send "{Esc}"
        Sleep 200
        capturedInfo.Push(A_Clipboard)
        Sleep 500
        if (gType = 1) {
            ; from Mainland
            ; capture: id
            MouseMove 738, 235
            Click "Down"
            Sleep 100
            MouseMove 483, 235
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: fullname
            MouseMove 658, 116
            Click "Down"
            Sleep 100
            MouseMove 498, 116
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: address
            MouseMove 750, 262
            Click "Down"
            Sleep 100
            MouseMove 498, 262
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100           
            capturedInfo.Push(A_Clipboard)
            Sleep 100
            ; capture: province
            ; TODO: add capture of province
        } else if (gType = 2) {
            ; from HK/MO/TW
            ; capture: id
            MouseMove 652, 291
            Click "Down"
            Sleep 100
            MouseMove 506, 291
            Click "Up"
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: fullname
            MouseMove 658, 116
            Click "Down"
            Sleep 100
            MouseMove 498, 116
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: lastname
            MouseMove 759, 203
            Click "Down"
            Sleep 100
            MouseMove 500, 203
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            MouseMove 759, 233
            Click "Down"
            Sleep 100
            MouseMove 500, 233
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            Sleep 100

        } else if (gType = 3) {
            ; from abroad
            ; capture: id
            MouseMove 652, 291
            Click "Down"
            Sleep 100
            MouseMove 759, 291
            Click "Up"
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: lastname
            MouseMove 759, 203
            Click "Down"
            Sleep 100
            MouseMove 500, 203
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: firstname
            MouseMove 759, 233
            Click "Down"
            Sleep 100
            MouseMove 500, 233
            Click "Up"
            Sleep 100
            Send "^c"
            Sleep 100
            capturedInfo.Push(A_Clipboard)
            ; capture: country
            MouseMove 537, 319
            Sleep 100
            Click 
            Sleep 100
            Click "Right"
            Sleep 100
            Send "c"
            Sleep 100
            Send "{Esc}"
            Sleep 200
            capturedInfo.Push(A_Clipboard)
            Sleep 100
        }
        WinSetAlwaysOnTop false, "旅客信息"
        return this.parseGuestInfo(gType, capturedInfo)
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
            if (StrLen(infoArr[3]) = 18) {
                guestProfile["idType"] := "IDC"
            } else if (StrLen(infoArr[3]) = 9) {
                guestProfile["idType"] := (SubStr(guestProfile["idNum"], 1, 1) = "C") ? "MRP" : "IDP"
            } else {
                guestProfile["idType"] := ""
            }
            guestProfile["address"] := infoArr[5]
            guestProfile["province"] := getProvince(infoArr[6])
        } else if (gType = 2) {
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
        } else if (gType = 3)  {
            ; from abroad
            guestProfile["language"] := "E"
            guestProfile["idType"] := "NOP"
            guestProfile["lastName"] := infoArr[4]
            guestProfile["firstName"] := infoArr[5]
            guestProfile["country"] :=  getCountryCode(infoArr[6])
        }

        ; debugging infos
        for k, v in guestProfile {
            toFill .= Format("{1}：{2}`n", k, v)
        }
        MsgBox(Format("
            (	
            即将填入的信息：

            {1}
            )", toFill), this.popupTitle, "4096")

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