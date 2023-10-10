class Phrases {
    static gbWidth := "w320 "
    
    static USE(App) {
        this.Rush(App)
        this.Upsell(App)
        this.TableReserve(App)
    }

    static Rush(App) {
        App.AddGroupBox(this.gbWidth . "r3", "RushRoom - 赶房与Key Keep")
        App.AddText("xp+10 yp+20 h20", "赶房时间")
        rushTime := App.AddEdit("x+10 w150 h20", "14:00")
        keyMade := [
            App.AddRadio("x35 yp+30 Checked", "已做卡"),
            App.AddRadio("x+10", "未做卡"),
        ]
        App.AddButton("xp w80 h50 x250 y195", "复制`nComment`nAlert").OnEvent("Click", copy)

        copy(*) {
            if (rushTime.Text = "" ) {
                return
            }
            A_Clipboard := (keyMade[1].Value = 1)
                ? Format("赶房 {1}, Key Keep L10", rushTime.Text)
                : Format("赶房 {1}, 未做房卡", rushTime.Text)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }

    static Upsell(App) {
        App.AddGroupBox(this.gbWidth . "r6 x25 y+20", "Upselling - 房间升级")
        App.AddText("xp+10 yp+25", "升级房型")
        roomType := App.AddEdit("x+10 w150", "")
        App.AddText("x35 y+10", "每晚差价")
        rateDiff := App.AddEdit("x+10 w150", "")
        App.AddText("x35 y+10", "升级晚数")
        nts := App.AddEdit("x+10 w150", "")
        App.AddText("x35 y+10", "语言版本")
        lang := [
            App.AddRadio("x+20 Checked", "中文"),
            App.AddRadio("x+10", "英文"),
        ]

        App.AddButton("xp w80 h50 x250 y285", "复制`nComment`nAlert").OnEvent("Click", copy)

        copy(*) {
            if (
                rateDiff.Text = "" ||
                roomType.Text = "" ||
                nts.Text = ""
            ) {
                return
            }
            A_Clipboard := (lang[1].Value = 1)
                ? Format("另加{1}元每晚 升级到{2}，共{3}晚，合共{4}元。", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
                : Format("Add RMB{1}(per night) upgrade to {2} for {3}Nights, total RMB{4}", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
            MsgBox(A_Clipboard, "已复制信息", "T1")
        }
    }

    static TableReserve(App) {
        App.AddGroupBox(this.gbWidth . "r7 x25 y+85", "TableResv - 餐饮预订")
        App.AddText("xp+10 yp+20", "预订餐厅")
        restaurant := App.AddComboBox("w150 x+10", ["宏图府", "玉堂春暖", "风味餐厅", "流浮阁"])
        App.AddText("x35 y+10", "预订日期")
        date := App.AddDateTime("w150 x+10", "LongDate")
        App.AddText("x35 y+10", "预订时间")
        time := App.AddEdit("x+10  w150", "")
        App.AddText("x35 y+10", "用餐人数")
        guests := App.AddEdit("x+10 w150", "")
        App.AddText("x35 y+10", "对方工号")
        staff := App.AddEdit("x+10 w150", "")

        App.AddButton("xp w80 h50 x250 y440", "复制`nComment`nAlert").OnEvent("Click", copy)

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