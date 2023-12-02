; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\lib\_JXON.ahk"
; #Include "%A_ScriptDir%\src\lib\RH_Macros\RH_Fedex.ahk"
; #Include "%A_ScriptDir%\src\lib\RH_Macros\RH_OTA.ahk"
; #Include "../../App.ahk"
#Include "../lib/utils.ahk"
#Include "../lib/_JXON.ahk"
#Include "./RH_Macros/RH_Fedex.ahk"
#Include "./RH_Macros/RH_OTA.ahk"

class ResvHandler {
    static name := 'Reservation Handler'
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static desc := ""

    resvTemp := IniRead(store, "ResvHandler", "ResvTemplates")
    static resvTempObj := Jxon_Load(&resvTemp)

    static USE(App) {
        OnClipboardChange(() => this.saveAddOnJson(App))

        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
            ; App.AddText("xp+10", this.desc),
            ; TODO: add template reservations and its setting here
            App.AddText("xp10 yp+10 h20", "奇利模板"),
            App.AddEdit("vkingsley x+10 h20", this.resvTempObj.kingsley),
            App.AddText("xp10 yp+10 h20", "Agoda模板"),
            App.AddEdit("vagoda x+10 h20", this.resvTempObj.agoda),
            App.AddButton("vstartBtn Default h35 w230 y+15", "开始录入预订"),
        ]
        
        startBtn := getCtrlByName("start", ui)
        kingsley := getCtrlByName("kingsley", ui)
        agoda := getCtrlByName("agoda", ui)
        tempEdits := getCtrlByTypeAll("Edit", ui)

        startBtn.OnEvent("Click", (*) => this.modifyReservation())
        for edit in tempEdits {
            edit.OnEvent("Change", (*) => saveTempConfirmation(edit))
        }

        saveTempConfirmation(curEdit) {
            curEditName := curEdit.Name
            ResvHandler.resvTempObj[curEditName] := Trim(curEdit.Text)
            IniWrite(Jxon_Dump(ResvHandler.resvTempObj), store, "ResvHandler", "ResvTemplates")
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
                App.Hide()
                ResvHandler.modifyReservation()
            } catch {
                MsgBox("请先打开 Opera 窗口。", ResvHandler.popupTitle)
            }
        }
    }

    static modifyReservation() {
        bookingInfo := IniRead(store, "ResvHandler", "JSON")
        bookingInfoObj := Jxon_Load(&bookingInfo)

        switch bookingInfoObj["agent"] {
            case "fedex":
                RH_Fedex(bookingInfoObj)
            case "kingsley":
                RH_Kingsley(bookingInfoObj, ResvHandler.resvTempObj.kingsley)
            default:
        }
        Sleep 500
    }
}