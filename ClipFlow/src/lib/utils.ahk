; getCtrlByName(vName, ctrlArray){
;     loop ctrlArray.Length {
;         if (vName = ctrlArray[A_Index].Name) {
;             return ctrlArray[A_Index]
;         }
;     }
; }

; getCtrlByType(ctrlType, ctrlArray){
;     loop  ctrlArray.Length {
;         if (type = ctrlArray[A_Index].Type) {
;             return ctrlArray[A_Index]
;         }
;     }
; }

; getCtrlByTypeAll(type, ctrlArray){
;     controls := []
;     loop ctrlArray.Length {
;         if (type = ctrlArray[A_Index].Type) {
;             controls.Push(ctrlArray[A_Index])
;         }
;     }
;     return controls
; }

; filePrepend(textToInsert, fileToPrepend){
;     textOrigin := FileRead(fileToPrepend)
;     FileDelete fileToPrepend
;     FileAppend textToInsert . textOrigin, fileToPrepend
; }


; cleanReload(quit := 0){
;     ; Windows set default
;     if (WinExist("ahk_class SunAwtFrame")) {
;         WinSetAlwaysOnTop false, "ahk_class SunAwtFrame"
;     }
;     if (WinExist("旅客信息")) {
;         WinSetAlwaysOnTop false, "旅客信息"
;     }
;     ; Key/Mouse state set default
;     BlockInput false
;     SetCapsLockState false
;     CoordMode "Mouse", "Screen"
;     if (quit = "quit") {
;         ExitApp
;     }
;     Reload
; }

; quitApp(*) {
;     quitConfirm := MsgBox("是否确定退出ClipFlow? ", popupTitle, "OKCancel 4096")
;     if (quitConfirm = "OK") {
;         cleanReload("quit")
;     } else {
;         cleanReload()
;     }
; }

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

    static arrayEvery(fn, targetArray) {
        if (!(fn is Func)) {
            throw TypeError(fn . " is not a Function Object.")
        }
        if (!(fn is Array)) {
            throw TypeError(targetArray . " is not an Array")
        }
        loop targetArray.Length {
            if (!fn(targetArray[A_Index])) {
                return false
            }
        }
        return true
    }

    static arrayFilter(fn, targetArray) {
        if (!(fn is Func)) {
            throw TypeError(fn . " is not a Function Object.")
        }
        if (!(fn is Array)) {
            throw TypeError(targetArray . " is not an Array")
        }
        newArray := []
        loop targetArray.Length {
            if (fn(targetArray[A_Index])) {
                newArray.Push(targetArray[A_Index])
            }
        }
        return newArray
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