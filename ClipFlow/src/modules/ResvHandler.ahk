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
        
    )"

    static USE(App) {
        OnClipboardChange this.saveAddOnJson
        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
            App.AddText("xp+10", this.desc),
            ; TODO: add template reservations and its setting here
            ; list view or just text+edit?
            App.AddButton("Default h35 w230 y+15", "开始录入预订"),
        ]
    }

    static saveAddOnJson() {
        if (!InStr(A_Clipboard, '"header":"RH"')) {
            return
        }
        IniWrite(A_Clipboard, store, "ResvHandler", "JSON")
        bookingInfoObj := Jxon_Load(&A_Clipboard)
        for k, v in bookingInfoObj {
            outputVal := IsObject(v) ? arrToStr(v) : v
            popupInfo .= Format("{1}：{2}`n", k, outputVal)
        }
        toOpera := MsgBox(Format("
        (   
        已获取订单的信息：

        {1}

        确定(Enter)：     打开 Opera
        取消(Esc)：       留在 当前页面
        )", popupInfo), this.popupTitle, "OKCancel 4096")
        if (toOpera = "OK") {
            try {
                WinActivate "ahk_class SunAwtFrame"
            } catch {
                MsgBox("请先打开 Opera 窗口。", this.popupTitle)
            }
        }
    }

    static parseReceivedInfo() {
        bookingInfo := IniRead(store, "ResvHandler", "JSON")
        bookingInfoObj := Jxon_Load(&bookingInfo)
        switch bookingInfoObj["agent"]{
            case "agoda":
                RH_Agoda(bookingInfoObj)
            default:
        }
    }
}