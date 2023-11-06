; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"
#Include "../../../App.ahk"

class ShareClip {
    static name := "Share Clip"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static shareClipFolder := "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\lib\SharedClips"
    static shareTxt := Format("{1}\{2}.txt", this.shareClipFolder, FormatTime(A_Now, "yyyyMMdd"))
    static prefix := Format("发送自: {1}, {2} `r`n", A_UserName, FormatTime(A_Now))
    static archiveDays := 5

    static USE(App) {
        OnClipboardChange this.listenAndSend
        ; create new txt
        if (!FileExist(this.shareTxt)) {
            FileCopy(A_ScriptDir . "\src\lib\shareTemplate.txt", this.shareTxt)
        }
        ; cleanup txts older than x days.
        loop files this.shareClipFolder "\*.txt" {
            if (DateDiff(FormatTime(A_Now, "yyyyMMdd"), SubStr(A_LoopFileName, 1, 6), "days") > this.archiveDays) {
                FileDelete this.shareClipFolder . "\" . A_LoopFileName
            }
        }
        ;states
        global isListening := false
        global readyToSend := false

        ui := [
            App.AddGroupBox("R15.5 w250 y+20", this.title),
            App.AddCheckbox("vclbListener " . "xp+10 yp+20 h15", "监听剪贴板"),
            App.AddText("xp y+15", "1、发送剪贴板History"),
            App.AddButton("vsendHistoryBtn " . "xp h32 w230 y+10", "发送 History"),
            App.AddText("xp y+15", "2、发送一段文字"),
            App.AddEdit("vuserInputText " . "xp h80 w230 y+10", ""),
            App.AddButton("vsendTextBtn" . "xp h32 w230 y+10", "发送 文字"),
            App.AddText("xp y+15", "3、查看 Share 剪贴板内容"),
            App.AddButton("vshowShareClipboardBtn" . "xp h32 w230 y+10", "打开 剪贴板"),
        ]
        ; get controls
        ; clbListener := ui[2]
        ; sendHistoryBtn := ui[4]
        ; userInputText := ui[6]
        ; sendTextBtn := ui[7]
        ; showShareClipboardBtn := ui[9]

        clbListener := getCtrlByName("clbListener" , ui)
        sendHistoryBtn := getCtrlByName("sendHistoryBtn", ui)
        userInputText := getCtrlByName("userInputText", ui)
        sendTextBtn := getCtrlByName("sendTextBtn", ui)
        showShareClipboardBtn := getCtrlByName("showShareClipboardBtn", ui)

        ; add events
        clbListener.OnEvent("Click", toggleListen)
        sendHistoryBtn.OnEvent("Click", sendHistory)
        userInputText.OnEvent("Change", checkInput)
        sendTextBtn.OnEvent("Click", sendUserInputText)
        showShareClipboardBtn.OnEvent("Click", showShareClipboard)

        ; callbacks
        toggleListen(*) {
            global isListening := !isListening
        }

        sendHistory(*) {
            clipHistory := strToArr(IniRead(store, "ClipHistory", "clipHisArr"))
            if (!clipHistory.Length) {
                return
            }
            text := this.sendHistory(clipHistory)
            MsgBox(Format("已发送：`r`n{1}", text), this.popupTitle, "4096 T1")
        }

        checkInput(*) {
            readyToSend := (userInputText.Text = "") ? false : true
        }

        sendUserInputText(*) {
            if (!readyToSend) {
                return
            }
            text := this.sendUserInputText(userInputText.Text)
            userInputText.Value := ""
            MsgBox(Format("已发送：`r`n{1}", text), this.popupTitle, "4096 T1")
        }

        showShareClipboard(*) {
            this.showShareClipboard(showShareClipboardBtn)
            showShareClipboardBtn.Enabled := false
        }
    }

    static listenAndSend() {
        this := ShareClip
        if (isListening = true) {
            FileAppend this.prefix . A_Clipboard . "`r`n`r`n", this.shareTxt
            MsgBox(this.prefix . A_Clipboard, this.popupTitle, "4096 T1")
        } else {
            return
        }
    }

    static sendHistory(clipHisArr) {
        text := this.prefix
        loop clipHisArr.Length {
            text .= clipHisArr[A_Index] . "`r`n"
        }
        FileAppend text . "`r`n`r`n", this.shareTxt
        return text
    }

    static sendUserInputText(userInput) {
        FileAppend this.prefix . userInput . "`r`n`r`n", this.shareTxt
        return this.prefix . userInput
    }

    static showShareClipboard(showShareClipboardBtn) {
        sharedText := FileRead(this.shareTxt)
        SetTimer(getUpdateSharedText, 2000)

        shareCLB := Gui(, "Share 剪贴板")
        ui := [
            shareCLB.AddEdit("w300 h480 ReadOnly", sharedText),
            shareCLB.AddButton("w100 h30 y+10", "打开源文件"),
        ]
        shareCLB.Show()
        WinSetAlwaysOnTop true, "Share 剪贴板"

        shareClipboard := ui[1]
        openShareClbBtn := ui[2]

        openShareClbBtn.OnEvent("Click", (*) => Run(this.shareTxt))
        shareCLB.OnEvent("Close", closeShareClipboardWin)

        getUpdateSharedText() {
            updatedText := FileRead(this.shareTxt)
            shareClipboard.Value := updatedText
        }

        closeShareClipboardWin(*) {
            showShareClipboardBtn.Enabled := true
            SetTimer(, 0)
        }
    }
}