; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"
#Include "../../../App.ahk"

class ShareClip {
    static name := "Share Clip"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static shareClipFolder := "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\lib\SharedClips"
    static shareTxt := Format("{1}\{2}.txt", this.shareClipFolder, FormatTime(A_Now, "yyyyMMdd"))

    static USE(App) {
        OnClipboardChange this.listenAndSend

        if (!FileExist(this.shareTxt)) {
            FileCopy(A_ScriptDir . "\src\lib\shareTemplate.txt", this.shareTxt)
        }

        ;states
        global isListening := true
        global readyToSend := false

        ui := [
            App.AddGroupBox("R15.5 w250 y+20", this.title),
            App.AddCheckbox("Checked xp+10 yp+20 h15", "监听剪贴板"),
            App.AddText("xp y+15", "1、发送剪贴板History"),
            App.AddButton("xp h32 w230 y+10", "发送 History"),
            App.AddText("xp y+15", "2、发送一段文字"),
            App.AddEdit("xp h80 w230 y+10", ""),
            App.AddButton("xp h32 w230 y+10", "发送 文字"),
            App.AddText("xp y+15", "3、查看 Share 剪贴板内容"),
            App.AddButton("xp h32 w230 y+10", "打开 剪贴板"),
        ]
        ; cleanup txts older than 5 days.
        loop files this.shareClipFolder "*.txt" {
            if (DateDiff(FormatTime(A_Now, "yyyyMMdd"), SubStr(A_LoopFileName, 1, 6), "days") >= 5) {
                FileDelete this.shareClipFolder . A_LoopFileName
            }
        }

        ; get controls
        clbListener := ui[2]
        sendHistoryBtn := ui[4]
        userInputText := ui[6]
        sendTextBtn := ui[7]
        showShareClipboardBtn := ui[9]

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

    static listenAndSend(){
        if (isListening = 1) {
            text := Format("发送自: {1}, {2} `r`n", A_UserName, FormatTime(A_Now))
            FileAppend text . A_Clipboard . "`r`n`r`n", ShareClip.shareTxt
            MsgBox(text . A_Clipboard, ShareClip.popupTitle, "4096 T1")
        } else {
            return
        }
    }

    static sendHistory(clipHisArr) {
        text := Format("发送自: {1}, {2} `r`n", A_UserName, FormatTime(A_Now))
        loop clipHisArr.Length {
            text .= clipHisArr[A_Index] . "`r`n"
        }
        FileAppend text . "`r`n`r`n", this.shareTxt
        return text
    }

    static sendUserInputText(userInput) {
        text := Format("发送自: {1}, {2} `r`n", A_UserName, FormatTime(A_Now))
        FileAppend text . userInput . "`r`n`r`n", this.shareTxt
        return text . userInput
    }

    static showShareClipboard(showShareClipboardBtn) {
        sharedText := FileRead(this.shareTxt)

        shareCLB := Gui(, "Share 剪贴板")
        ui := [
            shareCLB.AddEdit("w300 h480 ReadOnly", sharedText),
            shareCLB.AddButton("w100 h30 y+10", "打开源文件"),
        ]
        shareCLB.Show()

        openShareClbBtn := ui[2]

        openShareClbBtn.OnEvent("Click", (*) => Run(this.shareTxt))
        shareCLB.OnEvent("Close", closeShareClipboardWin)

        closeShareClipboardWin(*) {
            showShareClipboardBtn.Enabled := true
        }
    }
}