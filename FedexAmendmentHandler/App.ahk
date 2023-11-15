; #Include "%A_ScriptDir%\src\utils.ahk"
; #Include "%A_ScriptDir%\src\FedexAmendmentHandler.ahk"
#Include "../../FedexAmendmentHandler/src/utils.ahk"
#Include "../../FedexAmendmentHandler/src/FedexAmendmentHandler.ahk"

#SingleInstance Force
CoordMode "Mouse", "Screen"
; globals and states
version := "0.0.1"
popupTitle := "FedexAmendmentHandler " . version
templateFilePath := A_ScriptDir . "\src\FedExPasteTemplate.xls"

FAH := Gui("+Resize", popupTitle)
ui := [
    FAH.AddText("x20 y20", "1、 请将 FedEx 修改单粘贴到Excel模板"),
    FAH.AddEdit("vtemplatePath h25 w150 x20 y+10 ReadOnly", templateFilePath),
    FAH.AddButton("vshowTemplate h25 w70 x+20 Default", "打开Excel"),
    FAH.AddText("x20 y+20", "2、 如为修改单，请提供需要修改的订单确认号"),
    FAH.AddText("x20 y+10 h25", "确认号: "),
    FAH.AddEdit("vconfNum x+5 h25 w100 Number", ""),
    FAH.AddButton("vstartBtn x20 h25 w70 y+20", "开始修改"),
    FAH.AddButton("h25 w70 x+20", "退出脚本"),
]

FAH.Show()

; get ctrls
showTemplate := getCtrlByName("showTemplate", ui)
templatePath := getCtrlByName("templatePath", ui)
startBtn := getCtrlByName("startBtn", ui)
confNum := getCtrlByName("confNum", ui)

; add events
FAH.OnEvent("DropFiles", dropAmendment)
showTemplate.OnEvent("Click", openTemplate)
startBtn.OnEvent("Click", (*) => FedexAmendmentHandler.getAmendmentInfo())

; callbacks
dropAmendment(GuiObj, GuiCtrlObj, FileArray, X, Y, *) {
    if (FileArray.Length > 1) {
        FileArray.RemoveAt[1]
    }
    templatePath.Value := FileArray[1]
}

openTemplate(*){
	if (!InStr(templatePath.Text, "xls")) {
		MsgBox("请选择Excel表")
		return
	}
	Run(templateFilePath)
}