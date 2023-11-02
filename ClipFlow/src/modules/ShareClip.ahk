; #Include "%A_ScriptDir%\src\lib\utils.ahk"
#Include "../../lib/utils.ahk"
#Include "../../../App.ahk"

class ShareClip {
    static name := "Share Clip"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static shareTxt := Format("\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\lib\ShareClip-{1}.txt", FormatTime(A_Now, "yyyyMMdd"))
    static shareTxtMinusFive := Format("\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\src\lib\ShareClip-{1}.txt", FormatTime(DateAdd(A_Now, -5, "days"), "yyyyMMdd"))
    static clipHisArr := strToArr(IniRead(A_MyDocuments . "\ClipFlow.ini", "ClipHistory", "clipHisArr"))

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
        if (FileExist(this.shareTxtMinusFive)) {
            FileDelete this.shareTxtMinusFive
        }

        ; get controls
        sendHistoryBtn := ui[3]
        userInputText := ui[5]
        sendTextBtn := ui[6]
        showShareClipboardBtn := ui[8]

        ; add events
        sendHistoryBtn.OnEvent("Click", (*) => sendHistory(this.clipHisArr))
        sendTextBtn.OnEvent("Click", (*) => sendUserInputText(userInputText.Text))
        showShareClipboardBtn.OnEvent("Click", (*) => this.showShareClipboard)

        ; callbacks
        sendHistory(clipHisArr) {
            this.sendHistory(clipHisArr)
        }

        sendUserInputText(userInput) {
            this.sendUserInputText(userInput)
        }
    }

    static sendHistory(clipHisArr) {
        text := Format("发送自: {1}, {2}`n", A_UserName, FormatTime(A_Now))
        loop clipHisArr.Length {
            text .= clipHisArr[A_Index] . "`n"
        }
        FileAppend text . "`n`n", this.shareTxt
    }

    static sendUserInputText(userInput) {
        text := Format("发送自: {1}, {2}`n", A_UserName, FormatTime(A_Now))
        FileAppend text . userInput . "`n`n", this.shareTxt
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