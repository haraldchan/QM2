; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\pinyin.ahk"
#Include "../utils.ahk"
#Include "../DictIndex.ahk"

class PSB {
    static USE(App) {
        PSBinfo := "
        (
            Flow - Profile Mode
            
            1、请先打开 PSB 旅客界面，点击“开始复制”；
            2、复制完成后请打开Opera Profile 界面，
              点击“开始填入”。
        )"
        App.AddText("h20", PSBinfo)
        App.AddButton("Default h25 w80", "开始复制").OnEvent("Click", psbCopy)
        App.AddButton("h25 w80 x+20", "开始填入").OnEvent("Click", psbPaste)

        psbCopy(*) {
            App.Hide()
            Sleep 200
            global profileCache := PSB.Copy()
            App.Show()
        }
        
        psbPaste(*) {
            App.Hide()
            Sleep 200
            PSB.Paste(profileCache)
            App.Show()
            WinActivate "ahk_class SunAwtFrame"
        }
    }

    static popupTitle := "ClipFlow - PSB"

    static Copy() {
        ; guestType := 0
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
        ; TODO: try get guest type with getpixelcolor(get the radio that is selected), need more precise pixel
        ; CoordMode "Pixel", "Window"
        ; checkGuestType := [PixelGetColor(1,1), PixelGetColor(2,2), PixelGetColor(3,3)]
        ; loop checkGuestType.Length {
        ;     if (checkGuestType[A_Index] = "0x000000")
        ;         gType := A_Index
        ;         break
        ; }

        Sleep 500
        return this.Capture(guestType.Value)
        ; return this.Capture(gType)
    }
    
    static Capture(gType) {
        CoordMode "Mouse", "Window"
        BlockInput true
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
            Sleep 10
            MouseMove 483, 235
            Click "Up"
            Sleep 10
            Send "^c"
            Sleep 10
            capturedInfo.Push(A_Clipboard)
            ; capture: fullname
            MouseMove 658, 116
            Click "Down"
            Sleep 10
            MouseMove 498, 116
            Click "Up"
            Sleep 10
            Send "^c"
            Sleep 10
            capturedInfo.Push(A_Clipboard)
            ; capture: address
            MouseMove 750, 262
            Click "Down"
            Sleep 10
            MouseMove 498, 262
            Click "Up"
            Sleep 10
            Send "^c"
            Sleep 10           
            capturedInfo.Push(A_Clipboard)
            Sleep 10
            ; capture: province
            ; TODO: add capture of province
            MouseMove 587, 292
            Sleep 10
            Click 
            Sleep 10
            Click "Right"
            Sleep 10
            Send "c"
            Sleep 10
            Send "{Esc}"
            Sleep 20
            capturedInfo.Push(A_Clipboard)

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
            ; MouseMove 652, 291
            ; Click "Down"
            ; Sleep 100
            ; MouseMove 759, 291
            ; Click "Up"
            MouseMove 666, 290
            Sleep 100
            Click 2
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
            ; capture: country
            MouseMove 670, 322
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
        BlockInput false
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
            guestProfile["address"] := ""
            if (SubStr(guestProfile["idNum"], 1, 1) = "H") {
                guestProfile["idType"] := "HKC"
                guestProfile["province"] := "HK"
            } else if (SubStr(guestProfile["idNum"], 1, 1) = "M") {
                guestProfile["idType"] := "HKC"
                guestProfile["province"] := "MO"
            } else {
                guestProfile["idType"] := "TWT"
                guestProfile["province"] := "TW"
            }
        } else if (gType = 3)  {
            ; from abroad
            guestProfile["language"] := "E"
            guestProfile["idType"] := "NOP"
            guestProfile["lastName"] := infoArr[4]
            guestProfile["firstName"] := infoArr[5]
            guestProfile["country"] :=  getCountryCode(infoArr[6])
        }

        ; debugging popup
        for k, v in guestProfile {
            toFill .= Format("{1}：{2}`n", k, v)
        }
        MsgBox(Format("
            (   
            即将填入的信息：

            {1}
            )", toFill), this.popupTitle, "OKCancel 4096")

        return guestProfile
    }

    static Paste(guestProfileMap) {
        CoordMode "Mouse", "Screen"
        BlockInput true
        ; TODO:(advance feature) try set custom relative coords 
        ;       1.get profile window image
        ;       2.use ImageSearch to get initial(anchor) pixel pos
        ;       3.if it works, try implement it to other scripts

        ; TODO: mouse action to send(not paste) guest profile info
        ; { action: common info: lastName, firstName, language, gender, country, birthday, idType, idNum
        MouseMove 433, 286
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["lastName"])
        Sleep 10
        MouseMove 384, 313
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["firstName"])
        Sleep 10
        MouseMove 350, 341
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["language"])
        Sleep 10
        MouseMove 435, 339
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["gender"])
        Sleep 10
        MouseMove 346, 494
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["country"])
        Sleep 10
        MouseMove 843, 285
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["birthday"])
        Sleep 10
        MouseMove 829, 306
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["idType"])
        Sleep 100
        MouseMove 867, 326
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["idNum"])
        Sleep 100
        ; }

        if (guestProfileMap.Has("altName")) {
            ; action: with hanzi name
            ; address, province, altName, gender(in altName window)
        MouseMove 439, 370
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["address"])
        Sleep 10
        MouseMove 434, 492
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["province"])
        Sleep 10

        MouseMove 458, 282 ; open alt name win
        Click 1
        Sleep 4000
        MouseMove 712, 377
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["altName"])
        Sleep 10
        MouseMove 738, 464
        Click 3
        Sleep 10
        Send Format("{Text}{1}", guestProfileMap["gender"])
        Sleep 100
        }
        BlockInput false
    }
}