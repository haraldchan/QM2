; #Include "%A_ScriptDir%\src\lib\utils.ahk"
; #Include "%A_ScriptDir%\src\modules\RH_Reader.ahk"
#Include "../lib/utils.ahk"
#Include "../modules/RH_Reader.ahk"

class RsvnHandler {
    static name := "Reservation Handler"
    static title := "Flow Mode - Rsvn Handler"
    static popupTitle := "ClipFlow - Reservation Handler"
    static suptOTA := Map(
        "捷旅", RH_Jielv,
        "奇利", RH_Kingsley,
        "携程(现付)", RH_Ctrip,
        "携程(预付)", RH_CtripPrepaid,
        "飞猪(全场景/现付)", RH_Fliggy,
        "飞猪(核销券)", RH_FliggyGroup,
        "美团", RH_Meituan,
        "微信商城", RH_Wechat,
    )

    static USE(App) {
        otaList := []
        for k in this.suptOTA {
            otaList.Push(k)
        }

        ; GUI
        App.AddGroupBox("R6 w250", this.title)
        RH := App.AddTab3("xp+10 w240", ["使用步骤", "设置"])

        RH.UseTab(1)
        App.AddText("x20 h20", "1. 请先在 Ebooking 全选复制订单信息。")

        App.AddText("x20 h20", "2. 粘贴信息到中转Excel表")
        xlsPath := App.AddEdit("x20 h20", "")
        selectXls := App.AddButton("x+10 h20 w80", "选择文件")
        open := App.AddButton("x+10 h20 w80", "打开文件")

        App.AddText("x20 h20", "3. 选择来源 OTA，开始录入订单")
        chosenOTA := App.AddComboBox("w150 x20", otaList)
        modRsvn := App.AddButton("x+10 h20 w80", "录入订单")
        App.AddText("x20 h20", "4. 如有多个名字的 Party 房，请补完其他名字。")

        RH.UseTab(2)
        App.AddText("x20 y+30 ", "※ 清除Schedule Excel占用：`n（此功能会强制关闭所有Excel 文件，请先确保已保存修改）")
        App.AddButton("h25 w70 y+15", "清除Excel").OnEvent("Click", killExcel)

        ; functions
        xlsPath.OnEvent("LoseFocus", savePath)
        savePath(*) {
            IniWrite(xlsPath.Text, store, "RsvnHandler", "xlsPath")
        }

        selectXls.OnEvent("Click", getXls)
        getXls(*) {
            App.Opt("+OwnDialogs")
            selectFile := FileSelect(3, , "请选择 中转 文件")
            if (selectFile = "") {
                MsgBox("请选择文件")
                return
            }
            xlsPath.Value := selectFile
            IniWrite(selectFile, store, "RsvnHandler", "xlsPath")
        }

        open.OnEvent("Click", openXls)
        openXls(*) {
            Run xlsPath.Text
        }

        App.OnEvent("DropFiles", dropXls)
        dropXls(GuiObj, GuiCtrlObj, FileArray, X, Y, *) {
            if (FileArray.Length > 1) {
                FileArray.RemoveAt[1]
            }
            xlsPath.Value := FileArray[1]
            savePath()
        }

        killExcel(*) {
            if (ProcessExist("et.exe")) {
                ProcessClose "et.exe"
            } else if (ProcessExist("EXCEL.exe")) {
                ProcessClose "EXCEL.exe"
            } else {
                return
            }
        }
    }

    static ebookingHandler(ota) {
        ; ota pass from combobox
        Xl := ComObject("Excel.Application")
        source := Xl.Workbooks.Open(IniRead(store, "RsvnHandler", "xlsPath")).Worksheets(ota)

        ; call RH module of ota, which returns a ebooking info map, ready to fill into Opera
        ebookingInfoMap := this.suptOTA[ota](source)

        source.Close()
        XL.Quit()
        this.modifyRsvn(ebookingInfoMap)
    }

    static modifyRsvn(infoMap) {
        ; TODO: record actions for filling in fields.
        ; can borrow some actions from FSR maybe?
    }
}