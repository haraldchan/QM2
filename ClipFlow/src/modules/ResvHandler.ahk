; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../App.ahk"
#Include "../lib/utils.ahk"
#Include "../lib/_JXON.ahk"
#Include "./RH_Macros.ahk"

class ResvHandler {
    static name := 'Resv Handler'
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static desc := "
    (
        点击按键开始
    )"

    static USE(App) {
        OnClipboardChange this.saveAddOnJson
        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
            App.AddText("xp+10", this.desc),
            ; TODO: add template reservations and its setting here
            ; list view or just text+edit?
            App.AddButton("vstartBtn Default h35 w230 y+15", "开始录入预订"),
        ]

        startBtn := getCtrlByName("start", ui)

        startBtn.OnEvent("Click", this.modifyReservation)
    }

    static saveAddOnJson() {
        if (!InStr(A_Clipboard, '"header":"RH"')) {
            return
        }
        IniWrite(A_Clipboard, store, "ResvHandler", "JSON")
        clb := A_Clipboard
        bookingInfoObj := Jxon_Load(&clb)
        for k, v in bookingInfoObj {
            if (IsObject(v)) {
                if (k = "contacts") {
                    try { 
                        outputVal := "电话：" . v["phone"] . " " . "邮箱：" . v["email"] 
                        } catch {
                            outputVal := ""
                        }
                } else {
                    outputVal := arrToStr(v)
                }
            } else {
                outputVal := v
            }
            popupInfo .= Format("{1}：{2}`n", k, outputVal)
        }
        toOpera := MsgBox(Format("
        (   
        已获取订单的信息：

        {1}

        确定(Enter)：     打开 Opera
        取消(Esc)：       留在 当前页面
        )", popupInfo), ResvHandler.popupTitle, "OKCancel 4096")
        if (toOpera = "OK") {
            try {
                WinActivate "ahk_class SunAwtFrame"
            } catch {
                MsgBox("请先打开 Opera 窗口。", ResvHandler.popupTitle)
            }
        }
    }

    static modifyReservation() {
        bookingInfo := IniRead(store, "ResvHandler", "JSON")
        bookingInfoObj := Jxon_Load(&bookingInfo)
        switch bookingInfoObj["agent"]{
            case "fedex":
                RH_Fedex(bookingInfoObj)
            default:
        }
        Sleep 500
        MsgBox("已完成录入，请核对。", ResvHandler.popupTitle, "T3 4096")
    }
}