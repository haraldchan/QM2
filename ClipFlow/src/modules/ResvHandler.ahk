; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\lib\_JXON.ahk"
; #Include "%A_ScriptDir%\src\lib\RH_Macros\RH_Fedex.ahk"
; #Include "%A_ScriptDir%\src\lib\RH_Macros\RH_OTA.ahk"
; #Include "../../App.ahk"
#Include "../lib/utils.ahk"
#Include "../lib/_JXON.ahk"
#Include "./RH_Macros/RH_Fedex.ahk"
#Include "./RH_Macros/RH_OTA.ahk"

store := A_MyDocuments . "\ClipFlow.ini"
resvTemp := IniRead(store, "ResvHandler", "ResvTemplates")
resvTempObj := Jxon_Load(&resvTemp)

class ResvHandler {
    static name := "Reservation Handler"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static desc := ""
    static initXpos := 0
    static initYpos := 0

    static USE(App) {
        OnClipboardChange((*) => this.saveAddOnJson(App))
        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
            ; App.AddText("xp+10", this.desc),
            ; TODO: add template reservations and its setting here
            App.AddText("xp10 yp+20 h20", "奇利模板 "),
            App.AddEdit("vkingsley x+10 h20", resvTempObj["kingsley"]),
            App.AddText("x40 yp+30 h20", "Agoda模板"),
            App.AddEdit("vagoda x+10 h20", resvTempObj["agoda"]),
            App.AddButton("vstartBtn Default x40 h35 w230 y+15", "开始录入预订"),
        ]

        startBtn := Interface.getCtrlByName("startBtn", ui)
        kingsley := Interface.getCtrlByName("kingsley", ui)
        agoda := Interface.getCtrlByName("agoda", ui)
        tempEdits := Interface.getCtrlByTypeAll("Edit", ui)

        startBtn.OnEvent("Click", (*) => this.modifyReservation())
        loop tempEdits.Length {
            tempEdits[A_Index].OnEvent("Change", saveTempConfirmation.Bind(tempEdits[A_Index]))
        }

        saveTempConfirmation(curEdit,*) {
            curEditName := curEdit.Name
            resvTempObj[curEditName] := Trim(curEdit.Text)
            IniWrite(Jxon_Dump(resvTempObj), store, "ResvHandler", "ResvTemplates")
        }
    }

    static saveAddOnJson(App) {
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
                    outputVal := ""
                    loop v.Length {
                        outputVal .= v[A_Index] . "，"
                    }
                    outputVal := SubStr(outputVal, 1, StrLen(outputVal) - 1)
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
            ; try {
                WinActivate "ahk_class SunAwtFrame"
                App.Hide()
                ResvHandler.modifyReservation()
            ; } catch {
            ;     MsgBox("请先打开 Opera 窗口。", ResvHandler.popupTitle)
            ; }
        }
    }

    static modifyReservation() {
        bookingInfo := IniRead(store, "ResvHandler", "JSON")
        bookingInfoObj := Jxon_Load(&bookingInfo)

        switch bookingInfoObj["agent"] {
            case "fedex":
                RH_Fedex(bookingInfoObj)
            case "kingsley":
                RH_Kingsley(bookingInfoObj, resvTempObj["kingsley"])
            default:
        }
        Sleep 500
    }
}