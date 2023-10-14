; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"

class InvoiceWechat {
    static name := "Invoice Fill-in(Wechat)"
    static title := "Flow Mode - Invoice Fill-in(Wechat)"
    static popupTitle := "ClipFlow - Invoice Mode(Wechat)"
    static desc := "
    (
        Flow - Invoice Mode(Wechat)
        
        1、请先打开小程序“微信发票助手”界面，
          点击 “复制全部资料”。

        2、复制完成后请打开一键开票界面，点
         击“开始填入”。
    )"

    static USE(App) {
        OnClipboardChange this.parseInvoiceInfo
        ; GUI
        App.AddGroupBox("R6 w250 y+20", this.title)
        App.AddText("xp+10", this.desc)
        fillInBtn := App.AddButton("Default h35 w230 y+15", "开始填入")

        ; functions
        fillInBtn.OnEvent("Click", fillInfo)

        fillInfo(*){
            this.fillInfo(this.parseInvoiceInfo())
        }
    }

    static parseInvoiceInfo() {
        if (InStr(A_Clipboard, "名称：") && InStr(A_Clipboard, "税号：")) {
            invoiceInfo := StrSplit(A_Clipboard, "`n")
            ; MsgBox(invoiceInfo[1])
            invoiceInfoMap := Map()
            if (invoiceInfo.Length = 7) { ; probably is 6, check it
                invoiceInfoMap["company"] := SubStr(invoiceInfo[1], 4)
                invoiceInfoMap["taxNum"] := StrReplace(SubStr(invoiceInfo[2], 4), " ", "")
                invoiceInfoMap["address"] := SubStr(invoiceInfo[3], 6)
                invoiceInfoMap["tel"] := SubStr(invoiceInfo[4], 4)
                invoiceInfoMap["bank"] := SubStr(invoiceInfo[5], 6)
                invoiceInfoMap["account"] := SubStr(invoiceInfo[6], 6)
            } else {
                invoiceInfoMap["company"] := SubStr(invoiceInfo[1], 4)
                invoiceInfoMap["taxNum"] := StrReplace(SubStr(invoiceInfo[2], 4), " ", "")
            }
            try {
                WinActivate "ahk_exe VATIssue Terminal.exe"
            } catch {
                MsgBox("请先打开 一键开票。", InvoiceWechat.popupTitle)
            }
            return invoiceInfoMap
        } else {
            return
        }
    }

    static fillInfo(infoMap) {
        CoordMode "Mouse", "Screen"
        CoordMode "Pixel", "Screen"
        if (PixelGetColor(667, 163) = "0xFFFFCC") {
            MouseMove 735, 182
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1}", infoMap["company"])
            Sleep 10
            MouseMove 735, 214
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1}", infoMap["taxNum"])
            Sleep 10
            MouseMove 735, 242
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["address"], infoMap["tel"])
            Sleep 10
            MouseMove 735, 272
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["bank"], infoMap["account"])
            Sleep 10
        } else if (PixelGetColor(667, 163) = "0xCCFFFF") {
            MouseMove 730, 230
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1}", infoMap["company"])
            Sleep 10
            MouseMove 730, 263
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1}", infoMap["taxNum"])
            Sleep 10
            MouseMove 730, 292
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["address"], infoMap["tel"])
            Sleep 10
            MouseMove 730, 319
            Click
            Send "^a"
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["bank"], infoMap["account"])
            Sleep 10
        }
    }
}