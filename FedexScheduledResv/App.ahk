; { imports
; #Include "%A_ScriptDir%\utils.ahk"
; #Include "%A_ScriptDir%\FedexScheduledResv.ahk"
#Include "../utils.ahk"
#Include "../FedexScheduledResv.ahk"
; }

; { setup
#SingleInstance Force
TraySetIcon A_ScriptDir . "\assets\FSRTray.ico"
version := "2.0.0"
popupTitle := "Fedex Scheduled Reservations " . version
config := A_ScriptDir . "\config.ini"
path := IniRead(config, "FSR", "schedulePath")
resvOnDayTime := IniRead(config, "FSR", "resvOnDayTime")
try {
    WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
} catch {
    MsgBox("
    (
    当前为非管理员状态或未运行Opera PMS，将无法正确运行。
    请在FSR 图标上右键点击，选择“以管理员身份运行”；
    及打开Opera PMS。 
    )", "Fedex Scheduled Reservations")
    ExitApp
}
; }

; { GUI
FSR := Gui("+Resize", popupTitle)
FSR.AddText(, "
(
快捷键及对应功能：

Enter:     开始录入
F12:       强制停止脚本

)")
tab3 := FSR.AddTab3("w350", ["选项", "常见问题"])
tab3.UseTab(1)
FSR.AddText("x20 y+20", "1. 请选择Schedule 文件")
; schedule path (by input or file select)
schdPath := FSR.AddEdit("h25 w150 x20 y+10", path)
schdPath.OnEvent("LoseFocus", savePath)
savePath(*) {
    IniWrite(schdPath.Text, config, "FSR", "schedulePath")
}
FSR.AddButton("h25 w70 x+15", "选择文件").OnEvent("Click", getSchd)
getSchd(*) {
    FSR.Opt("+OwnDialogs")
    selectFile := FileSelect(3, , "请选择 Schedule 文件")
    if (selectFile = "") {
        MsgBox("请选择文件")
        return
    }
    schdPath.Value := selectFile
    IniWrite(selectFile, config, "FSR", "schedulePath")
}
FSR.AddButton("h25 w70 x+15", "打开文件").OnEvent("Click", openSchd)
openSchd(*) {
    Run schdPath.Text
}
FSR.AddText("x20 y+30 ", "2. 请指定提前一天留房时间点：`n（ETA在此时间点后的房间不会提前一天留房）")
toNextDropdown := FSR.AddDropDownList("w150 x20 y+10 Choose2", ["09:00", "10:00", "11:00", "12:00", "13:00"])
toNextDropdown.OnEvent("Change", saveTime)
saveTime(*) {
    IniWrite(toNextDropdown.Text, config, "FSR", "resvOnDayTime")
}

FSR.AddText("x20 y+30 ", "※ 清除Schedule Excel占用：`n（此功能会强制关闭所有Excel 文件，请先确保已保存修改）")
FSR.AddButton("h25 w70 y+15", "清除Excel").OnEvent("Click", killExcel)
killExcel(*) {
    if (ProcessExist("et.exe")) {
        ProcessClose "et.exe"
    } else if (ProcessExist("EXCEL.exe")) {
        ProcessClose "EXCEL.exe"
    } else {
        return
    }
}

tab3.UseTab(2)
FSR.AddText("x20 y+20", "
(
Q：发现运行错误后，我可以怎样停止脚本？
A：任何时候按下F12 即可停止脚本运行。

Q: 读入标签页弹窗应该填入什么？
A: 应填入标签页编号。比如读取"Sheet1"则填入"1"，
   读取"Sheet1-2"则填入"1-2"。

Q: 为什么脚本接管时会出现鼠标点击错误、或修改错误
   日期等导致信息填入错位？
A: 最大原因可能是IE 运行一段时候后占用内存过多导
   致Loading 时间过长所产生的错位。
   解决方法：关闭IE 浏览器，重新启动IE 重新登入
            Opera 即可。

Q：出现错误按下F12 终止后，我应该怎样继续使用脚本
   修改？
A：需要按如下步骤操作：

  1) 在“选项”页点击“清除Excel” 按钮。
    i. 注：此举会关闭所有excel 文件，如果有其他正
       在编辑的excel 文件请先保存！

  2) 打开Schedule.xls，选择已经完成修改的信息所在
     行并删除行。
    i. 如果通一个Trip 内已经完成了一间房的修改，将
       No. Of Rooms 从2 修改为1 即可。
    ii.建议创建副本标签页进行修改操作，避免影响原
       信息。
)")

; tab3.UseTab(3)
; FSR.AddPicture("w500 h-1", A_ScriptDir . "\assets\RsvnList.JPG")
tab3.UseTab()

FSR.AddButton("h40 w120 x40 Default", "开始录入").OnEvent("Click", runFSR)
runFSR(*) {
    if (schdPath.Text = "") {
        MsgBox("请选择 Schedule 文件")
        return
    }
    FSR.Hide()
    FsrMain()
}
FSR.AddButton("h40 w120 x+50", "退出脚本").OnEvent("Click", quit)
quit(*) {
    ExitApp
}

FSR.OnEvent("DropFiles", dropSchd)
dropSchd(GuiObj, GuiCtrlObj, FileArray, X, Y, *) {
    if (FileArray.Length > 1) {
        FileArray.RemoveAt[1]
    }
    schdPath.Value := FileArray[1]
    savePath()
}
FSR.Show()
; }

; hotkey setups
F12:: cleanReload()
#Hotif WinActive("ahk_class AutoHotkeyGUI")
Enter:: runFSR()
Esc:: quitApp()