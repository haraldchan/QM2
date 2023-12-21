#Include "../../../Lib/Classes/utils.ahk"

class Phrases {
    static gbWidth := "w320 "

    static USE(App) {
        this.Rush(App)
        this.Upsell(App)
        this.ExtraBed(App)
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
            App.AddEdit("vntsUps Number x+10 w150", ""),
            App.AddRadio("vlang h20 x35 y+10 Checked", "中文"),
            App.AddRadio("h20 x+10", "英文"),
            App.AddButton("xp w80 h50 x250 y285", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        roomType := Interface.getCtrlByName("roomType", ui)
        rateDiff := Interface.getCtrlByName("rateDiff", ui)
        ntsUps := Interface.getCtrlByName("ntsUps", ui)
        lang := Interface.getCtrlByName("lang", ui)

        copy(*) {
            if (
                rateDiff.Text = "" ||
                roomType.Text = "" ||
                ntsUps.Text = ""
            ) {
                return
            }
            A_Clipboard := (lang.Value = 1)
                ? Format("另加{1}元每晚 升级到{2}，共{3}晚，合共{4}元。", rateDiff.Text, roomType.Text, ntsUps.Text, rateDiff.Value * ntsUps.Value)
                : Format("Add RMB{1}(per night) upgrade to {2} for {3}Nights, total RMB{4}", rateDiff.Text, roomType.Text, ntsUps.Text, rateDiff.Value * ntsUps.Value)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }
    ; layout not tested yet
    static ExtraBed(App) {
        ui := [
            App.AddGroupBox(this.gbWidth . "r5 x25 y+80", "Extra Bed - 加床"),
            App.AddText("xp10 yp+25 r1", "价格"),
            App.AddRadio("Checked h15 x+10", "345元"),
            App.AddRadio("h15 x+5", "575元"),
            App.AddRadio("h15 x+5", "免费"),
            App.AddText("x35 h20 y+15", "批准人"),
            App.AddEdit("vapprover h20 w40 x+5", "Amy"),
            App.AddText("x+10", "加床晚数"),
            App.AddEdit("vntsEB Number x+10 w40", "1"),
            App.AddRadio("visChinese h20 x35 y+15 Checked", "中文"),
            App.AddRadio("h20 x+10", "英文"),
            App.AddButton("x250 y440 w80 h50", "复制`nComment`nAlert").OnEvent("Click", copy),
        ]

        payAmountOptions := [ui[3], ui[4], ui[5]]
        ntsEB := Interface.getCtrlByName("ntsEB", ui)
        approver := Interface.getCtrlByName("approver", ui)


        copy(*) {
            isChinese := Interface.getCtrlByName("isChinese", ui).value

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
                ? Format("另加床{1}元/晚，共{2}晚，合共{3}元。 ", selectedAmount, ntsEB.Value, selectedAmount * ntsEB.Value)
                : selectedAmount = 575
                    ? Format("另加床{1}元/晚含一位行政酒廊待遇，共{2}晚，合共{3}元。 ", selectedAmount, ntsEB.Value, selectedAmount * ntsEB.Value)
                    : Format("免费加床 by {1}，共{2}晚。 ", approver.Text, ntsEB.Value)
            commentEng := selectedAmount = 345
                ? Format("Add extra-bed for RMB{1}net per night for {2} night(s), RMB{3}net total. ", selectedAmount, ntsEB.Value, selectedAmount * ntsEB.Value)
                : selectedAmount = 575
                    ? Format("Add extra-bed for RMB{1}net per night(including ClubFloor benefits) for {2} night(s), RMB{3}net total.", selectedAmount, ntsEB.Value, selectedAmount * ntsEB.Value)
                    : Format("Free extra bed by {1} for {2}night(s). ", approver.Text, ntsEB.Value)

            A_Clipboard := isChinese ? commentChn : commentEng
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }

    static TableReserve(App) {
        ui := [
            App.AddGroupBox(this.gbWidth . "r7 x25 y+70", "Table Resv - 餐饮预订"),
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