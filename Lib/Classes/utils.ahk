#Include "./_JXON.ahk"

;Utils: general utility methods
class utils {
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
}

class json {
    static fileRead(jsonFile) {
        json := FileRead(jsonFile)
        return Jxon_Load(&json)
    }

    static fileWrite(newJson, jsonFile) {
        jsonOrigin := FileRead(jsonFile)
        FileDelete jsonOrigin
        FileAppend Jxon_Dump(newJson), jsonFile 
    }
}

; JSA: array methods that mimic JavaScript's Array.Prototype methods.
class jsa {
    static some(fn, targetArray) {
        if (!(fn is Func)) {
            throw TypeError("First parameter is not a Function Object.")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Second parameter is not an Array")
        }
        loop targetArray.Length {
            if (fn(targetArray[A_Index])) {
                return true
            }
        }
        return false
    }

    static every(fn, targetArray) {
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

    static filter(fn, targetArray) {
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

    static map(fn, targetArray) {
        if (!(fn is Func)) {
            throw TypeError("First parameter is not a Function Object.")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Second parameter is not an Array")
        }
        newArray := []
        loop targetArray.Length {
            newArray.Push(fn(targetArray[A_Index]))
        }
        return newArray
    }

    static reduce(fn, targetArray, initialValue := 0) {
        if (!(fn is Func)) {
            throw TypeError("First parameter is not a Function Object.")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Second parameter is not an Array")
        }
        if (!(initialValue is Number)) {
            throw TypeError("Third parameter is not an Number")
        }
        initIsSet := !(initialValue = 0)
        accumulator := initIsSet ? initialValue : targetArray[1]
        currentValue := initIsSet ? targetArray[1] : targetArray[2]
        loopTimes := initIsSet ? targetArray.Length : targetArray.Length - 1
        result := 0
        loop loopTimes {
            if (A_Index = 1) {
                result := fn(accumulator, currentValue)
            } else {
                if (!(initialValue = 0)) {
                    result := fn(result, targetArray[A_Index])
                } else {
                    result := fn(result, targetArray[A_Index + 1])
                }
            }
        }
        return result
    }

    static with(index, newValue, targetArray) {
        if (index > targetArray.Length) {
            throw ValueError("Index out of range")
        }
        if (!(targetArray is Array)) {
            throw TypeError("Third parameter is not an Array")
        }
        newArray := []
        loop targetArray.Length {
            newArray.Push(targetArray[A_Index])
        }
        newArray[index] := newValue
        return newArray
    }
}

; Interface: methods to interact with GUI controls.
class interface {
    static getCtrlByName(vName, ctrlArray) {
        loop ctrlArray.Length {
            if (ctrlArray[A_Index] is Array) {
                arrElement := ctrlArray[A_Index]
                this.getCtrlByName(vName, arrElement)
            }
            if (vName = ctrlArray[A_Index].Name) {
                return ctrlArray[A_Index]
            }
        }
    }

    static getCtrlByType(ctrlType, ctrlArray) {
        if (ctrlArray[A_Index] is Array) {
            arrElement := ctrlArray[A_Index]
            this.getCtrlByName(ctrlType, arrElement)
        }
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
 
    static getCheckedRowNumbers(listViewCtrl) {
		checkedRowNumbers := []
		loop listViewCtrl.GetCount() {
			curRow := listViewCtrl.GetNext(A_Index - 1, "Checked")
			try {
				if (curRow = prevRow || curRow = 0) {
					Continue
				}
			}
			checkedRowNumbers.Push(curRow)
			prevRow := curRow
		}
		return checkedRowNumbers
	}

	static getCheckedRowDataMap(listViewCtrl, mapKeys, checkedRows) {
		checkedRowsData := []
		for rowNumber in checkedRows {
			dataMap := Map()
			for key in mapKeys {
				dataMap[key] := listViewCtrl.GetText(rowNumber, A_Index)
			}
			checkedRowsData.Push(dataMap)
		}
		return checkedRowsData
	}
}