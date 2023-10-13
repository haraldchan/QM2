; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"

class InvoiceWechat {
    static name := "Invoice Fill-in(Wechat)"
    static title := "Flow Mode - Invoice Fill-in(Wechat)"
    static popupTitle := "ClipFlow - Invoice Mode(Wechat)"
    static invoiceCache := Map()
    static desc := "
    (
        Flow - Invoice Mode(Wechat)
        
        1、请先打开小程序“微信发票助手”界面，
          点击 “复制全部资料”。

        2、复制完成后请打开一键开票界面，点
         击“开始填入”。
    )"
    static invoiceAnchor := (A_OSVersion = "6.1.7601")
        ? "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\assets\invoiceAnchorWin7.PNG"
        : "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\assets\invoiceAnchor.PNG"

    static USE(App) {
        OnClipboardChange this.parseInvoiceInfo
        ; GUI
        App.AddGroupBox("R6 w250", this.title)
        App.AddText("xp+10", this.desc)
        fillInBtn := App.AddButton("Default h40 w200 y+15", "开始填入")

        ; functions
        fillInBtn.OnEvent("Click", this.fillInfo.Bind(InvoiceWechat, this.invoiceCache))
    }

    static parseInvoiceInfo() {
        if (InStr(A_Clipboard, "名称：") && InStr(A_Clipboard, "税号：")) {
            invoiceInfo := StrSplit(A_Clipboard, " `r`n")
            invoiceInfoMap := Map()
            if (invoiceInfo.Length = 7) { ; probably is 6, check it
                invoiceInfoMap["company"] := SubStr(invoiceInfo[1], 3)
                invoiceInfoMap["taxNum"] := StrReplace(SubStr(invoiceInfo[2], 3), " ", "")
                invoiceInfoMap["address"] := SubStr(invoiceInfo[3], 5)
                invoiceInfoMap["tel"] := SubStr(invoiceInfo[4], 3)
                invoiceInfoMap["bank"] := SubStr(invoiceInfo[5], 5)
                invoiceInfoMap["account"] := SubStr(invoiceInfo[6], 5)
                invoiceInfoMap["length"] := 7
            } else {
                invoiceInfoMap["company"] := SubStr(invoiceInfo[1], 3)
                invoiceInfoMap["taxNum"] := StrReplace(SubStr(invoiceInfo[2], 3), " ", "")
                invoiceInfoMap["length"] := 2
            }
            this.invoiceCache := invoiceInfoMap
            for k, v in invoiceInfoMap {
                popupInfo .= Format("{1}：{2}`n", k, v)
            }
            toIssue := MsgBox(Format("
                (   
                即将填入的信息：
    
                {1}
    
                确定(Enter)：     打开 Opera
                取消(Esc)：       留在 旅客信息
                )", popupInfo), this.popupTitle, "OKCancel")
            if (toIssue = "OK") {
                try {
                    ; TODO: change this to 一键开票
                    WinActivate "ahk_class SunAwtFrame"
                } catch {
                    MsgBox("请先打开 一键开票。", this.popupTitle)
                }
            }
        } else {
            return
        }
    }

    static fillInfo(infoMap) {
        CoordMode "Pixel", "Screen"
        CoordMode "Mouse", "Screen"
        if (ImageSearch(&FoundX, &FoundY, 0, 0, A_ScreenWidth, A_ScreenWidth, this.invoiceAnchor)) {
            anchorX := FoundX
            anchorY := FoundY
        } else {
            msgbox("请先打开开票界面", this.popupTitle)
            return
        }
        BlockInput true
        MouseMove anchorX + 0, anchorY + 0
        Click 3
        Sleep 10
        Send Format("{Text}{1}", infoMap["company"])
        Sleep 10
        MouseMove anchorX + 0, anchorY + 0
        Click 3
        Sleep 10
        Send Format("{Text}{1}", infoMap["taxNum"])
        Sleep 10

        if (infoMap["length"] > 2) {
            MouseMove anchorX + 0, anchorY + 0
            Click 3
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["address"], infoMap["tel"])
            Sleep 10
            MouseMove anchorX + 0, anchorY + 0
            Click 3
            Sleep 10
            Send Format("{Text}{1} {2}", infoMap["bank"], infoMap["account"])
            Sleep 10
        }
    }
}