#Include "../../../Lib/Classes/utils.ahk"

class Phrases {
    static gbWidth := "w320 "

    static USE(App) {
        this.Rush(App)
        this.Upsell(App)
        this.TableReserve(App)
    }

    static Rush(App) {
        ui := [
            App.AddGroupBox(this.gbWidth . "r3", "Rush Room - 赶房与Key Keep"),
            App.AddText("xp+10 yp+20 h20", "赶房时间"),
            App.AddEdit("vrushTime x+10 w150 h20", "14:00"),
            App.AddRadio("vkeyMade h25 x35 y+10 Checked", "已做卡"),
            App.AddRadio("h25 x+10", "未做卡"),
            App.AddButton("xp w80 h50 x250 y192", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        rushTime := Interface.getCtrlByName("rushTime", ui)
        keyMade := Interface.getCtrlByName("keyMade", ui)

        copy(*) {
            if (rushTime.Text = "") {
                return
            }
            A_Clipboard := (keyMade.Value = 1)
                ? Format("赶房 {1}, Key Keep L10", rushTime.Text)
                : Format("赶房 {1}, 未做房卡", rushTime.Text)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }

    static Upsell(App) {
        ui := [
            App.AddGroupBox(this.gbWidth . "r6 x25 y+20", "Upselling - 房间升级"),
            App.AddText("xp+10 yp+25", "升级房型"),
            App.AddEdit("vroomType x+10 w150", ""),
            App.AddText("x35 y+10 ", "每晚差价"),
            App.AddEdit("vrateDiff Number x+10 w150", ""),
            App.AddText("x35 y+10", "升级晚数"),
            App.AddEdit("vnts Number x+10 w150", ""),
            App.AddRadio("vlang h20 x35 y+10 Checked", "中文"),
            App.AddRadio("h20 x+10", "英文"),
            App.AddButton("xp w80 h50 x250 y285", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        roomType := Interface.getCtrlByName("roomType", ui)
        rateDiff := Interface.getCtrlByName("rateDiff", ui)
        nts := Interface.getCtrlByName("nts", ui)
        lang := Interface.getCtrlByName("lang", ui)

        copy(*) {
            if (
                rateDiff.Text = "" ||
                roomType.Text = "" ||
                nts.Text = ""
            ) {
                return
            }
            A_Clipboard := (lang.Value = 1)
                ? Format("另加{1}元每晚 升级到{2}，共{3}晚，合共{4}元。", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
                : Format("Add RMB{1}(per night) upgrade to {2} for {3}Nights, total RMB{4}", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }
    ; layout not tested yet
    static ExtraBed(App) {
        approverEnabled := false
        ui := [
            App.AddGroupBox(this.gbWidth . "r4 x25 y+20", "Extra Bed - 加床"),
            App.AddText("xp+10 yp+25", "加床金额"),
            [
                App.AddRadio("Default h25 x35 y+10", "345元"),
                App.AddRadio("h25 x+10", "575元"),
                App.AddRadio("h25 x+10", "免费"),
            ]
            App.AddText("h25 x+10", "批准人"),
            App.AddEdit("vapprover h25 w25 x+10", "")
            App.AddText("x35 y+10", "加床晚数"),
            App.AddEdit("vnts Number x+10 w150", "1"),
            App.AddRadio("visChinese h20 x35 y+10 Checked", "中文"),
            App.AddRadio("h20 x+10", "英文"),
            App.AddButton("xp w80 h50 x250 y285", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        payAmountOptions := ui[3]
        nts := Interface.getCtrlByName("nts", ui)
        approver := Interface.getCtrlByName("approver", ui)
        isChinese := Interface.getCtrlByName("isChinese", ui).value

        approver.Enabled := approverEnabled

        payAmountOptions[3].OnEvent("Click", toggleApprover)
        toggleApprover(*) {
            global approverEnabled := !approverEnabled
            approver.Enable := !approverEnabled
        }

        copy(*) {
            if (payAmountOptions[3].Value = 1 && approver.Text = "") {
                MsgBox("免费加床必须输入批准人", "常用语句", "T1")
                return
            }

            selectedAmount := payAmountOptions[1].Value = 1
                ? 345
                : payAmountOptions[2].Value = 1
                    ? 575
                    : 0

            commentChn := selectedAmount = 345
                ? Format("另加床{1}元/晚，共{2}晚，合共{3}元。 ", selectedAmount, nts.Value, selectedAmount * nts.Value)
                : selectedAmount = 575
                    ? Format("另加床{1}元/晚含一位行政酒廊待遇，共{2}晚，合共{3}元。 ", selectedAmount, nts, selectedAmount * nts)
                    : Format("免费加床 by {1}，共{2}晚。 ", approver.Text, nts.Value)


            commentEng := := selectedAmount = 345
                ? Format("Add extra-bed for RMB{1}net per night for {2} night(s), RMB{3}net total. ", selectedAmount, nts.Value, selectedAmount * nts.Value)
                : selectedAmount = 575
                    ? Format("Add extra-bed for RMB{1}net per night(including ClubFloor benefits) for {2} night(s), RMB{3}net total.", selectedAmount, nts, selectedAmount * nts)
                    : Format("Free extra bed by {1} for {2}night(s). ", approver.Text, nts.Value)

            A_Clipboard := isChinese ? commentChn : commentEng
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }

    static TableReserve(App) {
        ui := [
            App.AddGroupBox(this.gbWidth . "r7 x25 y+80", "Table Resv - 餐饮预订"),
            App.AddText("xp+10 yp+20", "预订餐厅"),
            App.AddComboBox("vrestaurant w150 x+10 Choose1", ["宏图府", "玉堂春暖", "风味餐厅", "流浮阁"]),
            App.AddText("x35 y+10", "预订日期"),
            App.AddDateTime("vdate w150 x+10", "LongDate"),
            App.AddText("x35 y+10", "预订时间"),
            App.AddEdit("vtime x+10  w150", ""),
            App.AddText("x35 y+10", "用餐人数"),
            App.AddEdit("vguests Number x+10 w150", ""),
            App.AddText("x35 y+10", "对方工号"),
            App.AddEdit("vstaff Number x+10 w150", ""),
            App.AddButton("xp w80 h50 x250 y440", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        restaurant := Interface.getCtrlByName("restaurant", ui)
        date := Interface.getCtrlByName("date", ui)
        time := Interface.getCtrlByName("time", ui)
        guests := Interface.getCtrlByName("guests", ui)
        staff := Interface.getCtrlByName("staff", ui)

        copy(*) {
            if (
                restaurant.Text = "" ||
                time.Text = "" ||
                guests.Text = "" ||
                staff.Text = ""
            ) {
                return
            }
            A_Clipboard := Format("已预订{1} {2}, {3} 人数: {4}位，接订工号：{5}",
                restaurant.Text,
                date.Text,
                time.Text,
                guests.Text,
                staff.Text)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }
}