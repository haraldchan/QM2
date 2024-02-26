#Include "../Lib/Classes/utils.ahk"
#Include "../Lib/Classes/Reactive.ahk"
#Include "../Lib/Classes/_JXON.ahk"
#Include "./src/modules/ProfileModify.ahk"
#Include "./src/modules/InvoiceWechat.ahk"
#Include "./src/modules/ShareClip.ahk"
#Include "./src/modules/ResvHandler.ahk"
#Include ../QM for FrontDesk/App.ahk

#SingleInstance Force
CoordMode "Mouse", "Screen"
TraySetIcon A_ScriptDir . "\src\assets\CFTray.ico"
OnClipboardChange addToHistory

version := "1.2.0"
popupTitle := "ClipFlow " . version
scriptHost := SubStr(A_ScriptDir, 1, InStr(A_ScriptDir, "\", , -1, -1) - 1)
modules := [
    ProfileModify,
    InvoiceWechat,
    ShareClip,
    ResvHandler,
]
winGroup := ["ahk_class SunAwtFrame", "旅客信息"]

iniChecker()
store := A_MyDocuments . "\ClipFlow.ini"
Sleep 100
tabPos := IniRead(store, "App", "tabPos")
clipHisObj := IniRead(store, "ClipHistory", "clipHisArr")
clipHisArr := Jxon_Load(&clipHisObj)
onTop := ReactiveSignal(false)

; GUI template
ClipFlow := Gui(, popupTitle)
ClipFlow.OnEvent("Close", (*) =>
    IniWrite(1, store, "App", "tabPos")
    utils.quitApp("ClipFlow", popupTitle, winGroup))
ClipFlow.AddCheckbox("h25 x15", "保持 ClipFlow 置顶    / 停止脚本: Ctrl+F12").OnEvent("Click", keepOnTop)

tab3 := ClipFlow.AddTab3("w280 x15 " . "Choose" . tabPos, ["Flow Modes", "History", "DevTool"])
tab3.OnEvent("Change", (*) => IniWrite(tab3.Value, store, "App", "tabPos"))

tab3.UseTab(1)
moduleLoader(ClipFlow)

tab3.UseTab(2)
tabHistory := []
renderHistory()

tab3.UseTab(3)
ClipFlow.AddText("x40 y80", "点击 Run Test 运行测试代码。")
ClipFlow.AddButton("y+20 h40 w160", "Run Test").OnEvent("Click", runTest)

tab3.UseTab()

ClipFlow.AddButton("h30 w130", "Clear").OnEvent("Click", clearList)
ClipFlow.AddButton("h30 w130 x+20", "Refresh").OnEvent("Click", (*) => utils.cleanReload(winGroup))

ClipFlow.Show()
; }

; { function scripts
keepOnTop(*) {
    onTop.set(onTop => !onTop)
    WinSetAlwaysOnTop onTop.get(), popupTitle
}

addToHistory(*) {
    if (A_Clipboard = "") {
        return
    }
    if (clipHisArr.Length = 10) {
        clipHisArr.Pop()
    }
    clipHisArr.InsertAt(1, A_Clipboard)
    IniWrite(Jxon_Dump(clipHisArr), store, "ClipHistory", "clipHisArr")
}

clearList(*) {
    FileDelete(store)
    FileCopy("../Lib/ClipFlow/ClipFlow.ini", A_MyDocuments)
    utils.cleanReload(winGroup)
}

iniChecker() {
    ; check ini file existence
    if (!FileExist(A_MyDocuments . "\ClipFlow.ini")) {
        FileCopy scriptHost . "\Lib\ClipFlow\ClipFlow.ini", A_MyDocuments
    }
    store := A_MyDocuments . "\ClipFlow.ini"
    ; check keys
    ini := Map(
        "App", ["tabPos", "moduleSelected"],
        "ClipHistory", "clipHisArr",
        "ResvHandler", "JSON"
    )
    try {
        for section, key in ini {
            if (key is Array) {
                for nestedKey in key {
                    IniRead(store, section, nestedKey)
                }
            } else {
                IniRead(store, section, key)
            }
        }
    } catch {
        FileDelete A_MyDocuments . "\ClipFlow.ini"
        Sleep 100
        FileCopy scriptHost . "\Lib\ClipFlow\ClipFlow.ini", A_MyDocuments
    }
}

moduleLoader(App) {
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

; render tab: Clipboard History base on clipHisArr
renderHistory() {
    global tabHistory := []
    if (clipHisArr.Length = 0) {
        return
    }
    loop clipHisArr.Length {
        tabHistory.Push(
            ClipFlow.AddEdit("x30 h40 w220 y+10 ReadOnly", clipHisArr[A_Index]),
            ClipFlow.AddButton("x+0 w30 h40", "×").OnEvent("Click", delHistoryItem.Bind(A_Index))
        )
    }

    delHistoryItem(index, *) {
        clipHisArr.RemoveAt(index)
        IniWrite(Jxon_Dump(clipHisArr), store, "ClipHistory", "clipHisArr")
        utils.cleanReload(winGroup)
    }
}

runTest(*) {
    ClipFlow.Hide()

    ; TEST CODE HERE

    utils.cleanReload(winGroup)
}

; hotkeys
^F12:: utils.cleanReload(winGroup)
Pause:: ClipFlow.Show()
#Hotif WinActive(popupTitle)
Esc:: ClipFlow.Hide()