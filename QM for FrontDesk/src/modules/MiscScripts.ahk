#Include "../../../Lib/Classes/utils.ahk"
#Include "../../../Lib/Classes/ReactiveControl.ahk"
CoordMode "Mouse", "Screen"
TrayTip "Private running"

class MiscScripts {
    static description := "杂项辅助脚本集（Billing、入Deposit等）"
    static popupTitle := "Misc Scripts"
    static pwd := ""
    static state := {
        passwordIsShow: false
    }

    static USE(){
        Misc := Gui("+Resize +AlwaysOnTop +MinSize250x300", this.popupTitle)
        ReactiveText(Misc, "h20", "Opera 密码")
        this.pwd := ReactiveEdit(Misc, "Password* h20 w100 x+10", "")
        ReactiveCheckbox(Misc, "h20 x+10", "显示", {event: "Click", callback: MiscScripts.togglePasswordVisibility})

        Misc.AddGroupBox("x10 w230 r5", "快捷键")
        Misc.AddText("xp+10 yp+20", "pw      - 快速输入密码")
        Misc.AddText("yp+20", "Alt+F11 - InHouse 界面打开Billing")

        Misc.Show()
    }

    static sendPassword(){
        userPwd := this.pwd.getInnerText()
        Send userPwd
    }

    static togglePasswordVisibility(){
        if (MiscScripts.state.passwordIsShow = true){
            MiscScripts.pwd.setOptions("+Password*")
        } else {
            MiscScripts.pwd.setOptions("-Password*") 
        }
        MiscScripts.state.passwordIsShow := !MiscScripts.state.passwordIsShow
    }

    static openBilling() {
        userPwd := this.pwd.getInnerText()
        try {
            WinMaximize "ahk_class SunAwtFrame"
            WinActivate "ahk_class SunAwtFrame"
        } catch {
            MsgBox("请先打开Opera PMS")
            return
        }
        Send "!t"
        Sleep 100
        Send "!b"
        Sleep 100
        Send Format("{Text}{1}", userPwd)
        Sleep 100
        Send "{Enter}"
    }
}

; funcs
; openBilling() {
;     try {
;         WinMaximize "ahk_class SunAwtFrame"
;         WinActivate "ahk_class SunAwtFrame"
;     } catch {
;         MsgBox("请先打开Opera PMS")
;         return
;     }
;     Send "!t"
;     Sleep 100
;     Send "!b"
;     Sleep 100
;     Send Format("{Text}{1}", pwd)
;     Sleep 100
;     Send "{Enter}"
; }

; deposit(paymentCode) {

;     try {
;         WinMaximize "ahk_class SunAwtFrame"
;         WinActivate "ahk_class SunAwtFrame"
;     } catch {
;         MsgBox("请先打开Opera PMS")
;         return
;     }
;     A_Clipboard := paymentCode
;     amount := InputBox("请输入金额")
;     supplement := InputBox("请输入单号（后四位即可）")
;     if (amount.Result = "Cancel") {
;         cleanReload()
;     }
;     if (supplement.Result = "Cancel") {
;         cleanReload()
;     }
;     Sleep 500
;     Send "!t"
;     MouseMove 710, 378
;     Sleep 500
;     Click 1
;     Sleep 500
;     Send "!p"
;     Sleep 200
;     MouseMove 584, 446
;     Sleep 300
;     Send "{BackSpace}"
;     Sleep 100
;     Send Format("{Text}{1}", pwd)
;     Sleep 200
;     MouseMove 707, 397
;     Sleep 500
;     Send "{Enter}"
;     Sleep 100
;     Send "^v"
;     Sleep 200
;     MouseMove 944, 387
;     Sleep 450
;     Send "{Tab}"
;     Sleep 200
;     Send "{Tab}"
;     Sleep 200
;     Send Format("{Text}{1}", amount.Value)
;     Sleep 200
;     MouseMove 577, 326
;     Sleep 100
;     Send "{Tab}"
;     Sleep 100
;     Send Format("{Text}{1}", supplement.Value)
;     Sleep 100
;     MouseMove 596, 421
;     Sleep 500
;     Send "!o"
;     Sleep 300
;     Send "{Escape}"
;     Sleep 200
;     Send "{Escape}"
;     Sleep 200
;     Send "{Escape}"
;     Sleep 200
;     Send "!c"
;     Sleep 200
;     Send "!c"
;     Sleep 200
; }

; blockBilling() {
;     try {
;         WinMaximize "ahk_class SunAwtFrame"
;         WinActivate "ahk_class SunAwtFrame"
;     } catch {
;         MsgBox("请先打开Opera PMS")
;         return
;     }
;     Send "{Enter}"
;     Sleep 100
;     Send "!r"
;     Sleep 100
;     Send "!a"
;     Sleep 100
;     loop 6{
;         Send "{Tab}"
;         Sleep 100
;     }
;     Send "{Text}-100"
;     Sleep 100
;     Send "!o"
;     MouseMove 695, 220
;     Click 1
;     Sleep 100
;     loop 10 {
;         Send "{PgDn}"
;         Sleep 10
;     }
;     Sleep 200
;     Send "!e"
;     loop 3 {
;         Send "{Enter}"
;         Sleep 100
;     }
;     Send "!t"
;     Sleep 100
;     Send "!b"
;     Sleep 100
;     Send Format("{Text}{1}", pwd)
;     Sleep 100
;     Send "{Enter}"
; }

; agoda8888(){
;     balance := InputBox("请输入账单金额")
;     orderId := InputBox("请输入单号")
;     Sleep 100
;     Send "!p"
;     Sleep 100
;     Send "8888"
;     Sleep 100
;     Send "{Tab}"
;     Send balance.Value
;     Sleep 100
;     loop 5 {
;         Send "{Tab}"
;         Sleep 10
;     }
;     Send Format("FR CRS.{1}", orderId.Value)
;     Sleep 100
;     Send "{Enter}"
;     Sleep 100
;     Send "8888"
;     Sleep 100
;     Send "{Tab}"
;     Sleep 100
;     Send Format("-{1}", balance.Value)
;     Sleep 100
;     loop 5 {
;         Send "{Tab}"
;         Sleep 10
;     }
;     Send "{Text}TO 9003"
;     Sleep 100
;     Send "{Enter}"
;     Sleep 100
;     Send "!o"
;     Sleep 100
;     Send "!c"
;     MsgBox("DONE.","Agoda Transfer", "T1 4096")
; }


; ########################
cleanReload(quit := 0){
    ; Windows set default
    if (WinExist("ahk_class SunAwtFrame")) {
        WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
    }
    ; Key/Mouse state set default
    BlockInput false
    SetCapsLockState false
    CoordMode "Mouse", "Screen"
    ; Excel Quit
    ComObject("Excel.Application").Quit()
    if (quit = "quit") {
        ExitApp
    }
    Reload
}

quitApp() {
    quitConfirm := MsgBox("是否确定退出Private? ", , "OKCancel 4096")
    if (quitConfirm = "OK") {
        cleanReload("quit")
    } else {
        cleanReload()
    }
}

; hotstrings setup
::rrr::{
    ; TrayTip "Reloaded"
    ; cleanReload()
}
::blk::{
    ; blockBilling()
}
::quit::{
    ; quitApp()
}
::agd::{
    ; agoda8888()
}
::pw::{
    ; Send Format("{Text}{1}", pwd)
}

; hotkeys setup
; #Hotif WinExist(MiscScripts.popupTitle)
::pw::{
    MiscScripts.sendPassword()
}
!F11::MiscScripts.openBilling()
; #F11::deposit("9132")