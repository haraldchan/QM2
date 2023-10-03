; #Include "%A_ScriptDir%\utils.ahk"
 #Include "../utils.ahk"

; { setup
#SingleInstance Force
TraySetIcon A_ScriptDir . "\assets\CFTray.ico"
version := "0.0.1"
popupTitle := "ClipFlow " . version
if (FileExist(A_MyDocuments . "\ClipFlow.ini")) {
    store := A_MyDocuments . "\ClipFlow.ini"
} else {
    FileCopy "\\10.0.2.13\fd\19-个人文件夹\HC\Software - 软件及脚本\AHK_Scripts\ClipFlow\ClipFlow.ini", A_MyDocuments
    store := A_MyDocuments . "\ClipFlow.ini"
}
clipHisArr := strToArr(IniRead(store, "clipHistory", "clipHistoryArr"))
flowArr := strToArr(IniRead(store, "Flow", "flowArr"))
isFlowCopying := false
isFlowPasting := false
flowPointer := 1
onTop := true
OnClipboardChange addToHistory
; }

; { GUI template
ClipFlow := Gui(, popupTitle)
ClipFlow.AddCheckbox("Checked h25 x20", "Keep On Top").OnEvent("Click", keepOnTop)
ClipFlow.AddButton("h20 w125", "Clear").OnEvent("Click", clearList)

tab3 := ClipFlow.AddTab3("w300 h700", ["ClipHistory", "Flow"])
tab3.UseTab(1)
tabFirst := []
renderHistory()

tab3.UseTab(2)
tabSecond := []
renderFlow()

tab3.UseTab()

ClipFlow.AddButton("h25 w100 x+18", "Flow Start").OnEvent("Click", flowStart)
ClipFlow.AddButton("h25 w100 x20", "Flow Load").OnEvent("Click", flowLoad)
; }

; { function scripts
addToHistory() {
    if (clipHisArr.Length > 10) {
        clipHisArr.Pop()
    }
    clipHisArr.InsertAt(1, A_Clipboard)
    renderHistory()
    IniWrite(arrToStr(clipHisArr), store, "ClipHistory", "clipHistoryArr")
}

keepOnTop(*) {
    global onTop := !onTop
    WinSetAlwaysOnTop !onTop, popupTitle
}

; render tab1: Clipboard History base on clipHisArr
renderHistory() {
    global tabFirst := []
    loop clipHisArr.Length {
        tabFirst.Push(
            ClipFlow.AddEdit("h50 w280 y+3 ReadOnly", clipHisArr[A_Index])
        )
    }
    loop clipHisArr.Length {
        tabFirst[A_Index].OnEvent("DoubleClick", setClb.Bind(tabFirst[A_Index]))
    }
    tab3.Redraw()
}

setClb(ctrlObj, *) {
    A_Clipboard := ctrlObj.Text
    MsgBox("Clipped!" popupTitle, "T1")
}

clearList(*) {
    clipHisArr := []
    renderHistory()
}

; flow related
; render tab2: Flow base on flowArr. render when complete FlowCoping
renderFlow() {
    loop flowArr.Length {
        tabSecond.Push(
            ClipFlow.AddEdit("h50 w280 y+3 ReadOnly", flowArr[A_Index])
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
            TrayTip "Flow loaded, ready to fire."
            IniWrite(arrToStr(flowArr), store, "Flow", flowArr)
        }
    }
}
; when flow turned on
flowAdd() {
    sleep 20
    flowArr.Push(1, clipHisArr[1])
}
; when flow turned off
flowLoad(*) {
    global isFlowCopying := false
    global isFlowPasting := true
    global tabSecond := []
    renderFlow()
    tab3.Redraw()
}

; over-ride default ctrl+v behavior
flowFire() {
    if (flowPointer > flowAdd.Length) {
        isFlowPasting := false
        global flowPointer := 1
        return
    }
    index := flowPointer
    Send Format("{Text}{1}", flowArr[index])
    global flowPointer++
}

; }
#HotIf isFlowCopying
^c:: flowAdd()
#HotIf isFlowPasting
~^v:: flowFire()
; double press Escape to ditch flow
~Esc::{
    if(KeyWait("Esc", "D T0.5")) {
        ditchClip := MsgBox("Ditch Flow Clip?", popupTitle, "OKCancel 4096")
        if (ditchClip = "OK") {
            global isFlowPasting := false
            global flowPointer := 1
        } else {
            return
        }
    }
}