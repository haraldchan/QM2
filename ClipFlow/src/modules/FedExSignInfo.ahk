#Include "../../../Lib/Classes/utils.ahk"
#Include "../../../FedexScheduledResv"

class FedExSignInfo {
    static name := "FedEx Sign-in Details"
    static title := "Flow Mode - " . this.name
    static popupTitle := "ClipFlow - " . this.name
    static desc := "
    (
        使用说明：

        1、打开需要获取信息的 FedEx修改单
          点击“获取信息”。

        2、成功复制信息后在 Sign-in表格中
          Name单元格粘贴即可。
    )"

    static USE() {

    }


}

class FedExGetInfo {
    static USE() {

    }

    ; from profile
    static getCrewName() {
        lastName := ""
        firstName := ""
        MouseMove 471, 217
        Click
        Sleep 3000
        ; get lastName
        MouseMove 380, 555
        Sleep 100
        Click 3
        Sleep 100
        Send "^c"
        Sleep 100
        lastName := A_Clipboard
        
        ; get firstName
        Send "{Tab}"
        Sleep 100
        Send "^c"
        Sleep 100
        firstName := A_Clipboard

        return Format("{1},`n{2}", lastName, firstName)
    }

    ; from More Fields
    static getDateTimeFlight() {
    
    }

    ; from TA.Rec log
    static getTripNumber(){
        
    }

    static getStayHours(ibDate, obDate, ETA, ETD){
        ; format in opera:
        ; MMddyyyy, hh:mm
        formattedIbDate := FormatTime(ibDate, "yyyyMMdd")
        formattedObDate := FormatTime(obDate, "yyyyMMdd")
        formattedETA := StrReplace(ETA, ":", "")
        formattedETD := StrReplace(ETD, ":", "")

        checkin := FormatTime(Format("{1}{2}", formattedIbDate, formattedETA))
        checkout := FormatTime(Format("{1}{2}", formattedObDate, formattedETD))
        stayMinutes := DateDiff(checkout, checkin, "Minutes")

        ;935min
        h := stayMinutes / 60
        m := stayMinutes - (60 * h)
        return Format("{1}:{2}", h, m)
    }
}