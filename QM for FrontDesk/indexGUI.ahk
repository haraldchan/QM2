#SingleInstance Force
QM := Gui("+Resize","QM for FrontDesk 2")
QM.AddText(, "manual texts`n")

tab3 := QM.AddTab3("w300 h200",["basics","Excel Dp","ReportMaster"])
tab3.UseTab(1)
radio1 := QM.AddRadio("h15 y+10", "InHouse Share")
radio2 := QM.AddRadio("h15 y+10", "PayBy PayFor")
radio3 := QM.AddRadio("h15 y+10", "GroupShareDnm")
radio4 := QM.AddRadio("h15 y+10", "DoNotMove Only")
QM.AddButton("h25 w70 x30 y195", "启动脚本")
QM.AddButton("h25 w70 x+20", "退出")

tab3.UseTab(2)
radio5 := QM.AddRadio("h15 y+10", "GroupKeys")
radio6 := QM.AddRadio("h15 y+10", "GroupProfilesModify")
radio7 := QM.AddRadio("h15 y+10", "PsbBatchCO")
QM.AddButton("h25 w70 x30 y195", "启动脚本")
QM.AddButton("h25 w70 x+20", "退出")

QM.Show()