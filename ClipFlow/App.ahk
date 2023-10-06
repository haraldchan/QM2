; #Include "%A_ScriptDir%\utils.ahk"
 #Include "../utils.ahk"
#Include "../ProfileModify.ahk"

; { setup
#SingleInstance Force
CoordMode "Mouse", "Screen"
TraySetIcon A_ScriptDir . "\assets\CFTray.ico"
version := "0.0.1"
popupTitle := "ClipFlow " . version
if (FileExist(A_MyDocuments . "\ClipFlow.ini")) {
    store := A_MyDocuments . "\ClipFlow.ini"
} else {
    FileCopy "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\ClipFlow.ini", A_MyDocuments
    store := A_MyDocuments . "\ClipFlow.ini"
}
clipHisArr := strToArr(IniRead(store, "ClipHistory", "clipHisArr"))
flowArr := strToArr(IniRead(store, "Flow", "flowArr"))
flowPointer := 1
isFlowCopying := false
isFlowPasting := false
onTop := true
OnClipboardChange addToHistory()
; }

; { GUI template
ClipFlow := Gui(, popupTitle)
ClipFlow.OnEvent("Close", quitApp)
ClipFlow.AddText(,"~~When Flowing, press Esc to Unflow~~")
ClipFlow.AddCheckbox("Checked h25 x20", "Keep ClipFlow On Top").OnEvent("Click", keepOnTop)
ClipFlow.AddButton("Disabled h25 w85", "Flow Start").OnEvent("Click", flowStart)
ClipFlow.AddButton("Disabled h25 w85 x+12", "Flow Load").OnEvent("Click", flowLoad)
ClipFlow.AddButton("Disabled h25 w85 x+12", "Load History").OnEvent("Click", loadAsFlow)

tab3 := ClipFlow.AddTab3("w280 x20", ["Flow Modes", "Flow", "History"])

tab3.UseTab(1)
PSBinfo := "
(
    Flow - Profile Mode
    
    1、请先打开 PSB 旅客界面，点击“开始复制”；
    2、复制完成后请打开Opera Profile 界面，
      点击“开始填入”。
)"
ClipFlow.AddText("h20", PSBinfo)
ClipFlow.AddButton("Default h25 w80", "开始复制").OnEvent("Click", psbCopy)
ClipFlow.AddButton("h25 w80 x+20", "开始填入").OnEvent("Click", psbPaste)
ProfileModify.USE(ClipFlow)

tab3.UseTab(2)
tabFlow := []
renderFlow()

tab3.UseTab(3)
tabHistory := []
renderHistory()

; TODO: learn reservation patterns, add it as ResvMode!

tab3.UseTab()

ClipFlow.AddButton("h25 w130", "Clear").OnEvent("Click", clearList)
ClipFlow.AddButton("h25 w130 x+20", "Refresh").OnEvent("Click", refresh)

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
    IniWrite(arrToStr(clipHisArr), store, "ClipHistory", "clipHisArr")
    if (isFlowCopying) {
        flowAdd()
    }
}

; render tab1: Clipboard History base on clipHisArr
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

clearList(*) {
    FileDelete(store)
    FileCopy("\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\ClipFlow.ini", A_MyDocuments)
    cleanReload()
}
refresh(*) {
    cleanReload()
}

; flow related
; render tab2: Flow base on flowArr. render when complete FlowCoping
renderFlow() {
    global flowArr := strToArr(IniRead(store, "Flow", "flowArr"))
    loop flowArr.Length {
        tabFlow.Push(
            ClipFlow.AddEdit("h30 w250 y+10 ReadOnly", flowArr[A_Index])
        )
    }
}

flowStart(*) {
    global isFlowCopying := !isFlowCopying
    bgc := isFlowCopying ? "AEE9FF" : ""
    ClipFlow.BackColor := bgc
    global flowArr := []
    if (isFlowCopying) {
        flowing := MsgBox("Flowing! OK to end flow.", popupTitle, "OK 4096")
        if (flowing = "OK") {
            flowLoad()
        }
    }
}
; when flow turned on
flowAdd() {
    sleep 20
    flowArr.Push(clipHisArr[1])
}
; when flow turned off
flowLoad(*) {
    global isFlowCopying := false
    global isFlowPasting := true
    A_Clipboard := ""
    MsgBox("Flow loaded, ready to fire.", popupTitle, "4096 T1")
    IniWrite(arrToStr(flowArr), store, "Flow", "flowArr")
}

; over-ride default ctrl+v behavior
flowFire() {
    if (flowPointer > flowArr.Length) {
        isFlowPasting := false
        global flowPointer := 1
    }
    Send flowArr[flowPointer]
    flowPointer++
}

loadAsFlow(*) {
    global flowArr := clipHisArr
    flowLoad()
}

; PSB to Opera
psbCopy(*) {
    ClipFlow.Hide()
    Sleep 200
    global profileCache := ProfileModify.Copy()
    ClipFlow.Show()
}

psbPaste(*) {
    ClipFlow.Hide()
    Sleep 200
    ProfileModify.Paste(profileCache)
    ClipFlow.Show()
    WinActivate "ahk_class SunAwtFrame"
}
; }

; hotkeys
Pause::{
    ClipFlow.Show()
}
#HotIf isFlowPasting
~^v:: flowFire()
; double press Escape to unload flow
~Esc::{
    if(KeyWait("Esc", "D T0.5")) {
        unload := MsgBox("Stop Flowing?", popupTitle, "OKCancel 4096")
        if (unload = "OK") {
            global isFlowPasting := false
            global flowPointer := 1
            ClipFlow.BackColor := ""
        } else {
            return
        }
    }
}