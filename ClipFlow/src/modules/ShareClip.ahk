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
        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
            App.AddText("xp+10", "1、发送剪贴板History"),
            App.AddButton("h35 w230 y+15", "发送 History"),
            App.AddText("xp+10 y+20", "2、发送一段文字"),
            App.AddEdit("h50 w230 y+15", ""),
            App.AddButton("h35 w230 y+10", "发送 文字"),
            App.AddText("xp+10 y+20", "3、查看 Share 剪贴板内容"),
            App.AddButton("h35 w230 y+10", "打开 剪贴板"),
        ]
        ; cleanup txts older than 5 days.
        loop files this.shareClipFolder "*.txt" {
            if (DateDiff(FormatTime(A_Now, "yyyyMMdd"), SubStr(A_LoopFileName, -4, 4), "days") >= 5) {
                FileDelete this.shareClipFolder . A_LoopFileName
            }
        }

        ; get controls
        sendHistoryBtn := ui[3]
        userInputText := ui[5]
        sendTextBtn := ui[6]
        showShareClipboardBtn := ui[8]

        ; add events
        sendHistoryBtn.OnEvent("Click", (*) => sendHistory)
        sendTextBtn.OnEvent("Click", (*) => sendUserInputText)
        showShareClipboardBtn.OnEvent("Click", (*) => this.showShareClipboard)

        ; callbacks
        sendHistory() {
            clipHistory := strToArr(IniRead(A_MyDocuments . "\ClipFlow.ini", "ClipHistory", "clipHisArr"))
            text := this.sendHistory(clipHistory)
            MsgBox(Format("已发送：`n`n{1}", text), this.popupTitle, "4096 T1")
        }
        
        sendUserInputText() {
            text := this.sendUserInputText(userInputText.Text)
            userInputText.Value := ""
            MsgBox(Format("已发送：`n`n{1}", text), this.popupTitle, "4096 T1")
        }
    }

    static sendHistory(clipHisArr) {
        text := Format("发送自: {1}, {2}`n", A_UserName, FormatTime(A_Now))
        loop clipHisArr.Length {
            text .= clipHisArr[A_Index] . "`n"
        }
        FileAppend text . "`n`n", this.shareTxt
        return text
    }

    static sendUserInputText(userInput) {
        text := Format("发送自: {1}, {2}`n", A_UserName, FormatTime(A_Now))
        FileAppend text . userInput . "`n`n", this.shareTxt
        return text . userInput
    }

    static showShareClipboard() {
        sharedText := FileRead(this.shareTxt)

        shareCLB := Gui(, "Share 剪贴板")
        ui := [
            shareCLB.AddEdit("w500 h600 ReadOnly", sharedText)
            shareCLB.AddButton("w80 h30 x420", "打开剪贴板文件")
        ]
        shareCLB.Show()

        openShareClbBtn := ui[2]

        openShareClbBtn.OnEvent("Click", (*) => Run(this.shareTxt))
    }
}