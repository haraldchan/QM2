class Phrases {
    static gbWidth := "w250"
    
    static USE(App) {
        this.Rush(App)
        this.Upsell(App)
        this.TableReserve(App)
    }

    static Rush(App) {
        App.AddGroupBox(this.gbWidth, "RushRoom - 赶房与Key Keep")
        App.AddText("xp+10", "赶房时间")
        rushTime := App.AddEdit("x+10 w50", "14:00")
        keyMade := [
            App.AddRadio("Checked", "已做卡")
            App.AddRadio("x+10", "未做卡")
        ]
        App.AddButton("w80", "复制Comment/Alert").OnEvent("Click", copy)

        copy(*) {
            A_Clipboard := (keyMade[1].Value = 1)
                ? Format("赶房 {1}, Key Keep L10", rushTime.Text)
                : Format("赶房 {1}, 未做房卡", rushTime.Text)
        }
    }

    static Upsell(App) {
        App.AddGroupBox(this.gbWidth, "Upselling - 房间升级")
        App.AddText("xp+10", "升级房型")
        roomType := App.AddEdit("x+10 w50", "")
        App.AddText("xp+10", "每晚差价")
        rateDiff := App.AddEdit("x+10 w50", "")
        App.AddText("xp+10", "升级晚数")
        nts := App.AddEdit("x+10 w50", "")
        App.AddText("xp+10", "语言版本")
        lang := [
            App.AddRadio("Checked", "中文")
            App.AddRadio("x+10", "英文")
        ]

        App.AddButton("w80", "复制Comment/Alert").OnEvent("Click", copy)

        copy(*) {
            A_Clipboard := (lang[1].Value = 1)
                ? Format("另加{1}元每晚 升级到{2}，共{3}晚，合共{4}元。", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
                : Format("Add RMB{1}(per night) upgrade to {2} for {3}Nights, total RMB{4}", rateDiff.Text, roomType.Text, nts.Text, rateDiff.Value * nts.Value)
        }
    }

    static TableReserve(App) {
        App.AddGroupBox(this.gbWidth, "TableResv - 餐饮预订")
        App.AddText("xp+10", "餐厅")
        restaurant := App.AddComboBox("x+10", ["宏图府", "玉堂春暖", "风味餐厅", "流浮阁"])
        App.AddText("xp+10", "日期")
        date := App.AddDateTime("x+10", "LongDate")
        App.AddText("xp+10", "时间")
        time := App.AddEdit("x+10 w50", "")
        App.AddText("xp+10", "人数")
        guests := App.AddEdit("x+10 w50", "")
        App.AddText("xp+10", "工号")
        staff := App.AddEdit("x+10 w50", "")

        App.AddButton("w80", "复制Comment/Alert").OnEvent("Click", copy)

        copy(*) {
            A_Clipboard := Format("已预订{1} {2},{3} {4}位，接订工号：{5}", restaurant.Text, date.Value, time.Text, guests.Text, staff.Text)
        }
    }
}