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
modules := [
    ProfileModify,
    InvoiceWechat, 
    ShareClip,  
    ResvHandler,
]
winGroup := ["ahk_class SunAwtFrame", "旅客信息"]
version := "1.1.0"
scriptHost := "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\QM2 - Nightly"
popupTitle := "ClipFlow " . version
if (!FileExist(A_MyDocuments . "\ClipFlow.ini")) {
    FileCopy scriptHost . "\Lib\ClipFlow\ClipFlow.ini", A_MyDocuments
}
store := A_MyDocuments . "\ClipFlow.ini"
Sleep 100
tabPos := IniRead(store, "App", "tabPos")
clipHisObj := IniRead(store, "ClipHistory", "clipHisArr")
clipHisArr := Jxon_Load(&clipHisObj)
onTop := false


; GUI template
ClipFlow := Gui(, popupTitle)
ClipFlow.OnEvent("Close", (*) => 
    IniWrite(1, store, "App", "tabPos")
    Utils.quitApp("ClipFlow", popupTitle, winGroup)
    )
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
ClipFlow.AddButton("h30 w130 x+20", "Refresh").OnEvent("Click", (*) => Utils.cleanReload(winGroup))

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

moduleLoader(App) {
    moduleSelected := IniRead(store, "App", "moduleSelected")
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
                IniWrite(A_Index, store, "App", "moduleSelected")
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
            ClipFlow.AddEdit("x30 h40 w220 y+10 ReadOnly", clipHisArr[A_Index]),
            ClipFlow.AddButton("x+0 w30 h40","×").OnEvent("Click", delHistoryItem.Bind(A_Index))
        )
    }

    delHistoryItem(index, *){
        clipHisArr.RemoveAt(index)
        IniWrite(Jxon_Dump(clipHisArr), store, "ClipHistory", "clipHisArr")
        Utils.cleanReload(winGroup)
    }
}

runTest(*) {
    ClipFlow.Hide()

    ; test code here.

    Utils.cleanReload(winGroup)
}

; hotkeys
^F12:: Utils.cleanReload(winGroup)
Pause:: ClipFlow.Show()
#Hotif WinActive(popupTitle)
Esc:: ClipFlow.Hide()