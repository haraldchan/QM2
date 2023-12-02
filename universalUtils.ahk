; winGroup := ["ahk_class SunAwtFrame", "旅客信息"]

class Utils {
    static cleanReload(winGroup, quit := 0) {
        ; Windows set default
        loop winGroup.Length {
            if (WinExist(winGroup[A_Index])) {
                WinSetAlwaysOnTop false, winGroup[A_Index]
            }
        }
        ; Key/Mouse state set default
        BlockInput false
        SetCapsLockState false
        CoordMode "Mouse", "Screen"
        if (quit = "quit") {
            ExitApp
        }
        Reload
    }

    static quitApp(appName, popupTitle, winGroup) {
        quitConfirm := MsgBox(Format("是否退出 {1}？", appName), popupTitle, "OKCancel 4096")
        quitConfirm = "OK" ? this.cleanReload(winGroup, "quit") : this.cleanReload(winGroup)
    }

    static filePrepend(textToInsert, fileToPrepend) {
        textOrigin := FileRead(fileToPrepend)
        FileDelete fileToPrepend
        FileAppend textToInsert . textOrigin, fileToPrepend
    }
}

class Interface {
    getCtrlByName(vName, ctrlArray) {
        loop ctrlArray.Length {
            if (vName = ctrlArray[A_Index].Name) {
                return ctrlArray[A_Index]
            }
        }
    }
    
    getCtrlByType(ctrlType, ctrlArray) {
        loop ctrlArray.Length {
            if (type = ctrlArray[A_Index].Type) {
                return ctrlArray[A_Index]
            }
        }
    }
    
    getCtrlByTypeAll(type, ctrlArray) {
        controls := []
        loop ctrlArray.Length {
            if (type = ctrlArray[A_Index].Type) {
                controls.Push(ctrlArray[A_Index])
            }
        }
        return controls
    }    
}