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

    static mouseMoveOffset(initX, initY, targetX, targetY) {
        offsetX := targetX - initX
        offsetY := targetY - initY
        MouseMove initX + offsetX, initY + offsetY
    }

    static arrayEvery(fn, targetArray) {
        if (!(fn is Func)) {
            throw TypeError("First parameter is not a Function Object.")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Second parameter is not an Array")
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
            throw TypeError("First parameter is not a Function Object.")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Second parameter is not an Array")
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
    static getCtrlByName(vName, ctrlArray) {
        loop ctrlArray.Length {
            if (vName = ctrlArray[A_Index].Name) {
                return ctrlArray[A_Index]
            }
        }
    }

    static getCtrlByType(ctrlType, ctrlArray) {
        loop ctrlArray.Length {
            if (type = ctrlArray[A_Index].Type) {
                return ctrlArray[A_Index]
            }
        }
    }

    static getCtrlByTypeAll(type, ctrlArray) {
        controls := []
        loop ctrlArray.Length {
            if (type = ctrlArray[A_Index].Type) {
                controls.Push(ctrlArray[A_Index])
            }
        }
        return controls
    }
}