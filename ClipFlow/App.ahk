; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\modules\ProfileModify.ahk"
; #Include "./src/lib/utils.ahk"
#Include "../Lib/Classes/utils.ahk"
#Include "../Lib/Classes/_JXON.ahk"
#Include "./src/modules/ProfileModify.ahk"
#Include "./src/modules/InvoiceWechat.ahk"
#Include "./src/modules/ShareClip.ahk"
#Include "./src/modules/ResvHandler.ahk"
; { setup
#SingleInstance Force
CoordMode "Mouse", "Screen"
TraySetIcon A_ScriptDir . "\src\assets\CFTray.ico"
OnClipboardChange addToHistory
modules := [ProfileModify, InvoiceWechat, ShareClip, ResvHandler]
winGroup := ["ahk_class SunAwtFrame", "旅客信息"]
version := "DEV"
popupTitle := "ClipFlow " . version
if (FileExist(A_MyDocuments . "\ClipFlow.ini")) {
    store := A_MyDocuments . "\ClipFlow.ini"
} else {
    ; FileCopy A_ScriptDir . "\src\lib\ClipFlow.ini", A_MyDocuments
    FileCopy "../Lib/ClipFlow/ClipFlow.ini", A_MyDocuments
    store := A_MyDocuments . "\ClipFlow.ini"
}
Sleep 100
clipHisObj := IniRead(store, "ClipHistory", "clipHisArr")
clipHisArr := Jxon_Load(&clipHisObj)
onTop := false
; }

; { GUI template
ClipFlow := Gui(, popupTitle)
ClipFlow.OnEvent("Close", (*) => Utils.quitApp("ClipFlow", popupTitle, winGroup))
ClipFlow.AddCheckbox("h25 x15", "Keep ClipFlow On Top").OnEvent("Click", keepOnTop)

tab3 := ClipFlow.AddTab3("w280 x15", ["Flow Modes", "History", "DevTool"])

tab3.UseTab(1)
moduleLoader(ClipFlow)

tab3.UseTab(2)
tabHistory := []
renderHistory()

tab3.UseTab(3)
ClipFlow.AddButton("h30 w130", "Run Test").OnEvent("Click", runTest)

tab3.UseTab()

ClipFlow.AddButton("h30 w130", "Clear").OnEvent("Click", clearList)
ClipFlow.AddButton("h30 w130 x+20", "Refresh").OnEvent("Click", refresh)

ClipFlow.Show()
WinSetAlwaysOnTop onTop, popupTitle
; }

; { function scripts
keepOnTop(*) {
    global onTop := !onTop
    WinSetAlwaysOnTop onTop, popupTitle
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
    ; FileCopy(A_ScriptDir . "\src\lib\ClipFlow.ini", A_MyDocuments)
    FileCopy("../Lib/ClipFlow/ClipFlow.ini", A_MyDocuments)
    Utils.cleanReload(winGroup)
}

refresh(*) {
    Utils.cleanReload(winGroup)
}

moduleLoader(App) {
    moduleSelected := IniRead(store, "Module", "moduleSelected")
    loadedModules := []
    ; create module select radio
    loop modules.Length {
        moduleRadioStyle := (A_Index = moduleSelected) ? "h15 x30 y+10 Checked" : "h15 x30 y+10"
        loadedModules.Push(App.AddRadio(moduleRadioStyle, modules[A_Index].name))
    }
    ; add event
    loop loadedModules.Length {
        loadedModules[A_Index].OnEvent("Click", saveSelect)
    }
    ; load selected module
    modules[moduleSelected].USE(App)

    ; check which module is selected, save it to ini, reload (swap to seleced module)
    saveSelect(*) {
        loop loadedModules.Length {
            if (loadedModules[A_Index].Value = 1) {
                IniWrite(A_Index, store, "Module", "ModuleSelected")
                Utils.cleanReload(winGroup)
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
            ClipFlow.AddEdit("h30 w250 y+10 ReadOnly", clipHisArr[A_Index])
        )
    }
}

runTest(*) {
    ClipFlow.Hide()
    ; Entry.profileEntry(["chen","ming hao","陈明颢"])
    ; sleep 1000
    ; Entry.roomQtyEntry(5)
    ; sleep 1000
    ; Entry.dateTimeEntry("20241201","20241203")
    ; sleep 1000
    ; Entry.commentOrderIdEntry("1103619", "RM INCL 2BBF TO TA")
    sleep 1000
    bbf := [1, 1, 1]
    ; msgbox(bbf is Array)
    if (!(Utils.arrayEvery((item) => item = 0, bbf))) {
        Entry.breakfastEntry(bbf)
    }
    Utils.cleanReload(winGroup)
}

; hotkeys
F12:: Utils.cleanReload(winGroup)
Pause:: ClipFlow.Show()
#Hotif WinActive(popupTitle)
Esc:: ClipFlow.Hide()