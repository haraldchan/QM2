#Include "../Lib/Classes/utils.ahk"
#Include "../Lib/Classes/_JXON.ahk"
#Include "./src/modules/ProfileModify.ahk"
#Include "./src/modules/InvoiceWechat.ahk"
#Include "./src/modules/ShareClip.ahk"
#Include "./src/modules/ResvHandler.ahk"

#SingleInstance Force
CoordMode "Mouse", "Screen"
TraySetIcon A_ScriptDir . "\src\assets\CFTray.ico"
OnClipboardChange addToHistory

winGroup := ["ahk_class SunAwtFrame", "旅客信息"]
scriptHost := SubStr(A_ScriptDir, 1, InStr(A_ScriptDir, "\", , -1, -1) - 1)
if (!FileExist(A_MyDocuments . "\ClipFlow.ini")) {
    FileCopy scriptHost . "\Lib\ClipFlow\ClipFlow.ini", A_MyDocuments
}
store := A_MyDocuments . "\ClipFlow.ini"

class ClipFlowApp {
    static version := "1.3.0"
    static popupTitle := "ClipFlow " . version
    static tabPos := IniRead(store, "App", "tabPos")
    static clipHisObj := IniRead(store, "ClipHistory", "clipHisArr")
    static clipHisArr := Jxon_Load(&clipHisObj)
    static state := {
        onTop: false,
        App: {}
    }
    static modules := [
        ProfileModify,
        InvoiceWechat,
        ShareClip,
        ResvHandler,
    ]

    static USE() {
        OnClipboardChange(addToHistory)

        ClipFlow := Gui(, popupTitle)
        ClipFlow.OnEvent("Close", (*) =>
            IniWrite(1, store, "App", "tabPos")
            utils.quitApp("ClipFlow", this.popupTitle, winGroup))
        ClipFlow.AddCheckbox("h25 x15", "保持 ClipFlow 置顶    / 停止脚本: Ctrl+F12").OnEvent("Click", keepOnTop)

        tab3 := ClipFlow.AddTab3("w280 x15 " . "Choose" . tabPos, ["Flow Modes", "History", "DevTool"])
        tab3.OnEvent("Change", (*) => IniWrite(tab3.Value, store, "App", "tabPos"))

        tab3.UseTab(1)
        this.moduleLoader(ClipFlow)

        tab3.UseTab(2)
        historyEdits := []
        loop 10 {
            ClipFlow.AddEdit("x30 h30 w220 y+10 ReadOnly", "")
        }
        for edit in historyEdits {
            try {
                edit.Value := this.clipHisArr[A_Index]
            } catch {
                edit.Value := ""
            }
        }

        tab3.UseTab(3)
        ClipFlow.AddText("x40 y80", "点击 Run Test 运行测试代码。")
        ClipFlow.AddButton("y+20 h40 w160", "Run Test").OnEvent("Click", (*) => this.runTest())

        tab3.UseTab()

        ClipFlow.AddButton("h30 w130", "Clear").OnEvent("Click", clearList)
        ClipFlow.AddButton("h30 w130 x+20", "Refresh").OnEvent("Click", (*) => utils.cleanReload(winGroup))

        ClipFlow.Show()
        this.state.App := ClipFlow

        keepOnTop(*) {
            this.state.onTop := !this.state.onTop
            WinSetAlwaysOnTop this.state.onTop, this.popupTitle
        }

        clearList(*) {
            FileDelete(store)
            FileCopy("../Lib/ClipFlow/ClipFlow.ini", A_MyDocuments)
            utils.cleanReload(winGroup)
        }
    }

    static moduleLoader(App) {
        moduleSelectedStored := IniRead(store, "App", "moduleSelected")
        moduleSelected := moduleSelectedStored > modules.Length ? 1 : moduleSelectedStored

        loadedModules := []
        ; create module select radio
        loop modules.Length {
            moduleRadioStyle := (A_Index = moduleSelected) ? "h15 x30 y+10 Checked" : "h15 x30 y+10"
            loadedModules.Push(App.AddRadio(moduleRadioStyle, modules[A_Index].name))
        }
        ; add event
        for control in loadedModules {
            control.OnEvent("Click", saveSelect)
        }
        ; load selected module
        modules[moduleSelected].USE(App)

        ; check which module is selected, save it to ini, reload (swap to seleced module)
        saveSelect(*) {
            loop loadedModules.Length {
                if (loadedModules[A_Index].Value = 1) {
                    IniWrite(A_Index, store, "App", "moduleSelected")
                    utils.cleanReload(winGroup)
                }
            }
        }

    }

    static addToHistory() {
        if (A_Clipboard = "") {
            return
        }
        if (this.clipHisArr.Length = 10) {
            this.clipHisArr.Pop()
        }
        this.clipHisArr.InsertAt(1, A_Clipboard)
        IniWrite(Jxon_Dump(this.clipHisArr), store, "ClipHistory", "clipHisArr")
    }

    static runTest() {
        ClipFlow.Hide()

        ; TEST CODE HERE

        utils.cleanReload(winGroup)
    }
}

ClipFlowApp.USE()

; hotkeys
^F12:: utils.cleanReload(winGroup)
Pause:: ClipFlowApp.state.App.Show()
#Hotif WinActive(ClipFlowApp.popupTitle)
Esc:: ClipFlowApp.state.App.Show()