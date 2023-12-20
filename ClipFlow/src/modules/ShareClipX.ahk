#Include "../../../Lib/Classes/utils.ahk"
#Include "../../../Lib/Classes/_JXON.ahk"
#Include "../../App.ahk"
#Include ../../../FedexScheduledResv/App.ahk

class ShareClipX {
    static name := "Share Clip X"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static scriptHost := SubStr(A_ScriptDir, 1, InStr(A_ScriptDir, "\", , -1, -1) - 1)
    static tempDir := this.scriptHost . "\Lib\ClipFlow\temp"
    static SavedClips := Gui(, "Share 剪贴板")
    static SavedClipsUI := []

    static USE(App) {
        ui := [
            App.AddGroupBox("R6 w250 y+20", this.title),
        ]
    }

    static showSavedClips() {
        this.SavedClips.Show()
    }

    static checkFileType(content) {
        imageExtensions := ["jpg", "jpeg", "bmp", "png", "gif"]
        fileExtensions := ["ahk", "txt", "csv", "xls", "xlsx", "doc", "docx", "ppt", "pptx", "json", "ini", "html", "pdf", "js"]

        if (!InStr(content, ".")) {
            return "isText"
        } else {
            extension := StrSplit(StrLower(content), ".").Pop()
            if (JSA.some((item) => item = extension, imageExtensions)) {
                return "isImage"
            } else if (JSA.some((item) => item == extension, fileExtensions)) {
                return "isFile"
            } else {
                return "isText"
            }
        }
    }

    static onCopy() {
        clippedType := this.checkFileType(A_Clipboard)

        if (clippedType == "isText") {
            this.insertText()
        } else if (clippedType == "isImage") {
            this.insertImage()
        } else if (clippedType == "isFile") {
            this.insertFile()
        } else {
            return
        }
        MsgBox(Format("已复制：`n`n{1}", A_Clipboard), this.popupTitle, "4096 T1")
    }

    static insertText() {
        destPath := Format("{1}\#####{2}#####", this.tempDir, A_Now . ".txt")
        FileAppend(A_Clipboard, destPath)
        this.SavedClipsUI.InsertAt(1,
            [
                this.SavedClips.AddGroupBox("r3 w250 y+20", "文字 - Text"),
                this.SavedClips.AddEdit("xp+10 h30 w200", A_Clipboard),
                this.SavedClips.AddButton("h15 w15 x+0", "×").OnEvent("Click", (*) => ShareClipX.removeCLipFile(destPath)),
                this.SavedClips.AddButton("h15 w15 y+15", "✎").OnEvent("Click", (*) => ShareClipX.openClipFile(destPath)),
            ]
        )
    }

    static insertImage() {
        destPath := Format("{1}\#####{2}#####{3}", this.tempDir, A_Now, StrSplit(A_Clipboard, "\").Pop())
        FileMove(A_Clipboard, destPath)
        this.SavedClipsUI.InsertAt(1,
            [
                this.SavedClips.AddGroupBox("r5 w250 y+20", "图片 - Image"),
                this.SavedClips.AddPicture("xp h50 w200", destPath)
                this.SavedClips.AddButton("h15 w15 x+0", "×").OnEvent("Click", (*) => ShareClipX.removeCLipFile(destPath)),
                this.SavedClips.AddButton("h15 w15 y+15", "✎").OnEvent("Click", (*) => ShareClipX.openClipFile(destPath)),
            ]
        )
    }

    static insertFile() {
        fileName := StrSplit(A_Clipboard, "\").Pop()
        destPath := Format("{1}\#####{2}#####{3}", this.tempDir, A_Now, fileName)
        FileMove(A_Clipboard, destPath)
        this.SavedClipsUI.insertAt(1,
            [
                this.SavedClips.AddGroupBox("r5 w250 y+20", "文件 - File"),
                this.SavedClips.AddEdit("xp+10 h30 w200", fileName),
                this.SavedClips.AddButton("h15 w15 x+0", "×").OnEvent("Click", (*) => ShareClipX.removeCLipFile(destPath)),
                this.SavedClips.AddButton("h15 w15 y+15", "✎").OnEvent("Click", (*) => ShareClipX.openClipFile(destPath)),
            ]
        )
    }

    static removeCLipFile(file) {
        this := ShareClipX
        loop files this.tempDir {
            if (InStr(file, A_LoopFileName)) {
                FileDelete file
            }
        }
        Utils.cleanReload([winGroup])
    }

    static openClipFile(file) {
        this := ShareClipX
        loop files this.tempDir {
            if (InStr(file, A_LoopFileName)) {
                Run(file)
            }
        }
    }
}