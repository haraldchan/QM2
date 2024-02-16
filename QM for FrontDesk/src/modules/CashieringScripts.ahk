#Include "../../../Lib/Classes/utils.ahk"
#Include "../../../Lib/Classes/Reactive.ahk"
CoordMode "Mouse", "Screen"
TrayTip "Private running"

class CashieringScripts {
    static description := "杂项辅助脚本集（快速打开Billing、入Deposit等）"
    static popupTitle := "Cashiering Scripts"
    static pwd := ""
    static paymentType := Map(
        9002, "9002 - Bank Transfer",
        9132, "9132 - Alipay",
        9138, "9138 - EFT-Wechat"
    )
    static passwordIsShow := ReactiveSignal(false)

    static USE() {
        Misc := Gui("+AlwaysOnTop +MinSize250x300", this.popupTitle)
        addReactiveText(Misc, "h20", "Opera 密码")
        this.pwd := addReactiveEdit(Misc, "Password* h20 w110 x+10", "")
        addReactiveCheckBox(Misc, "h20 x+10", "显示", { event: "Click", callback: CashieringScripts.togglePasswordVisibility })

        ; persist hotkey shotcuts
        Misc.AddGroupBox("r4 x10 w240", "快捷键")
        Misc.AddText("xp+10 yp+20", "pw      - 快速输入密码")
        Misc.AddText("yp+20", "agd     - 生成Agoda BalanceTransfer")
        Misc.AddText("yp+20", "blk     - Block 界面打开PM (主管权限)")
        Misc.AddText("yp+20", "Alt+F11 - InHouse 界面打开Billing")

        Misc.AddGroupBox("r4 x10 y+20 w240", "Deposit")
        Misc.AddText("xp+10 yp+20 h20", "支付类型")
        paymentType := addReactiveComboBox(Misc, "yp+20 w200 Choose2", this.paymentType)
        depositBtn := addReactiveButton(Misc, "y+10 w80", "录入 &Deposit", { event: "Click", callback: (*) => this.depositEntry(paymentType.getValue())})

        Misc.Show()
    }

    static sendPassword() {
        Send Format("{Text}{1}", this.pwd.getInnerText())
    }

    static togglePasswordVisibility() {
        if (CashieringScripts.passwordIsShow.get() = true) {
            CashieringScripts.pwd.setOptions("+Password*")
        } else {
            CashieringScripts.pwd.setOptions("-Password*")
        }
        CashieringScripts.passwordIsShow.set(!CashieringScripts.passwordIsShow.get())
    }

    static openBilling() {
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
        Send Format("{Text}{1}", this.pwd.getInnerText())
        Sleep 100
        Send "{Enter}"
    }

    static depositEntry(paymentTypeSelected) {
        try {
            WinMaximize "ahk_class SunAwtFrame"
            WinActivate "ahk_class SunAwtFrame"
        } catch {
            MsgBox("请先打开Opera PMS")
            return
        }
        amount := InputBox("请输入金额")
        supplement := InputBox("请输入单号（后四位即可）")
        if (amount.Result = "Cancel") {
            utils.cleanReload(winGroup)
        }
        if (supplement.Result = "Cancel") {
            utils.cleanReload(winGroup)
        }
        Sleep 500
        Send "!t"
        MouseMove 710, 378
        Sleep 500
        Click 1
        Sleep 500
        Send "!p"
        Sleep 200
        MouseMove 584, 446
        Sleep 300
        Send "{BackSpace}"
        Sleep 100
        Send Format("{Text}{1}", this.pwd.getInnerText())
        Sleep 200
        MouseMove 707, 397
        Sleep 500
        Send "{Enter}"
        Sleep 100
        Send Format("{Text}{1}", paymentTypeSelected)
        Sleep 200
        MouseMove 944, 387
        Sleep 450
        Send "{Tab}"
        Sleep 200
        Send "{Tab}"
        Sleep 200
        Send Format("{Text}{1}", amount.Value)
        Sleep 200
        MouseMove 577, 326
        Sleep 100
        Send "{Tab}"
        Sleep 100
        Send Format("{Text}{1}", supplement.Value)
        Sleep 100
        MouseMove 596, 421
        Sleep 500
        Send "!o"
        Sleep 300
        Send "{Escape}"
        Sleep 200
        Send "{Escape}"
        Sleep 200
        Send "{Escape}"
        Sleep 200
        Send "!c"
        Sleep 200
        Send "!c"
        Sleep 200
    }

    static agodaBalanceTransfer(){
        balance := InputBox("请输入账单金额")
        orderId := InputBox("请输入单号")
        WinActivate "ahk_class SunAwtFrame"
        Sleep 100
        Send "!p"
        Sleep 100
        Send "8888"
        Sleep 100
        Send "{Tab}"
        Send balance.Value
        Sleep 100
        loop 5 {
            Send "{Tab}"
            Sleep 10
        }
        Send Format("FR CRS.{1}", orderId.Value)
        Sleep 100
        Send "{Enter}"
        Sleep 100
        Send "8888"
        Sleep 100
        Send "{Tab}"
        Sleep 100
        Send Format("-{1}", balance.Value)
        Sleep 100
        loop 5 {
            Send "{Tab}"
            Sleep 10
        }
        Send "{Text}TO 9003"
        Sleep 100
        Send "{Enter}"
        Sleep 100
        Send "!o"
        Sleep 100
        Send "!c"
        MsgBox("DONE.","Agoda Transfer", "T1 4096")
    }

    static blockPmBilling(){
        try {
            WinMaximize "ahk_class SunAwtFrame"
            WinActivate "ahk_class SunAwtFrame"
        } catch {
            MsgBox("请先打开Opera PMS")
            return
        }
        Send "{Enter}"
        Sleep 100
        Send "!r"
        Sleep 100
        Send "!a"
        Sleep 100
        loop 6{
            Send "{Tab}"
            Sleep 100
        }
        Send "{Text}-100"
        Sleep 100
        Send "!o"
        MouseMove 695, 220
        Click 1
        Sleep 100
        loop 10 {
            Send "{PgDn}"
            Sleep 10
        }
        Sleep 200
        Send "!e"
        loop 3 {
            Send "{Enter}"
            Sleep 100
        }
        Send "!t"
        Sleep 100
        Send "!b"
        Sleep 100
        Send Format("{Text}{1}", this.pwd.getInnerText())
        Sleep 100
        Send "{Enter}"
    }
    
}

::pw:: {
    CashieringScripts.sendPassword()
}
::agd::{
    CashieringScripts.agodaBalanceTransfer()
}
::blk::{
    CashieringScripts.blockPmBilling()
}
!F11:: CashieringScripts.openBilling()