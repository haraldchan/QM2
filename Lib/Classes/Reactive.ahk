class ReactiveSignal {
    __New(val) {
        this.val := val
        this.subs := []
        this.comps := []
    }

    get(mutateFunction := 0) {
        if (mutateFunction != 0 && mutateFunction is Func) {
            return mutateFunction(this.val)
        } else {
            return this.val
        }
    }

    set(newSignalValue) {
        if (newSignalValue = this.val) {
            return
        }
        ; update val with new value
        this.val := newSignalValue is Func
            ? newSignalValue(this.val)
            : newSignalValue
        for ctrl in this.subs {
            ctrl.update(this.val)
        }
        for comp in this.comps {
            comp.sync(this.val)
        }
    }

    addSub(controlInstance) {
        this.subs.Push(controlInstance)
    }

    addComp(computed) {
        this.comps.Push(computed)
    }
}

class ComputedSignal {
    __New(signal, mutation) {
        this.signal := signal
        this.mutation := mutation
        this.val := this.mutation(this.signal.get())
        this.subs := []

        signal.addComp(this)
    }

    sync(newVal) {
        for ctrl in this.subs {
            ctrl.update(newVal)
        }
    }

    addSub(controlInstance) {
        this.subs.Push(controlInstance)
    }
}

class ReactiveControl {
    __New(controlType, depend, GuiObject, options, innerText, event := 0) {
        this.depend := depend
        this.GuiObject := GuiObject
        this.options := options
        this.innerText := innerText
        this.ctrlType := controlType

        this.ctrl := this.GuiObject.Add(this.ctrlType, options, innerText)
        if (event != 0) {
            this.event := event[1]
            this.callback := event[2]
            this.ctrl.OnEvent(this.event, (*) => this.callback())
        }

        depend.addSub(this)
    }

    setOptions(newOptions) {
        this.ctrl.Opt(newOptions)
    }

    getInnerText() {
        return this.ctrl.Text
    }

    setInnerText(newInnerText) {
        this.ctrl.Text := newInnerText is Func
            ? newInnerText(this.ctrl.Text)
            : newInnerText
    }

    setEvent(event, callback) {
        this.ctrl.OnEvent(event, (*) => callback())
    }

    update(newValue) {
        if (this.ctrl is Gui.Text || this.ctrl is Gui.Button) {
            this.ctrl.Text := this.reformat(this.innerText, newValue)
        } else if (this.ctrl is Gui.Edit) {
            this.ctrl.Value := this.reformat(this.innerText, newValue)
        }
    }

    reformat(text, vals*) {
        newStr := ""
        ; reconcat(text, val) {
        ;     insertIndex := InStr(text, "{{")
        ;     return Format(SubStr(text, 0, insertIndex) . "{1}" . SubStr(text, insertIndex + 5), val)
        ; }

        reconcat(text, val) {
            lIndex := InStr(text, "{{")
            lPart := SubStr(text, 0, lIndex)

            rIndex := InStr(text, "}}")
            rPart := SubStr(text, rIndex + 1)

            return Format(lPart . "{1}" . rPart, val)
        }

        ; e.g:
        ; 0: "this is a text, val1 is {{ num1 }}, val2 is {{ num2 }}"
        ; 1: "this is a text, val1 is num1, val2 is {{ num2 }}"
        ; 2: "this is a text, val1 is num1, val2 is num2"
        newStr := text
        loop vals.Length {
            newStr := reconcat(newStr, vals[A_Index])
        }
        return newStr
    }
}

class addReactiveButton extends ReactiveControl {
    __New(depend, GuiObject, options, innerText, event := 0) {
        super.__New("Button", depend, GuiObject, options, innerText, event)
    }
}

class addReactiveEdit extends ReactiveControl {
    __New(depend, GuiObject, options, innerText, event := 0) {
        super.__New("Edit", depend, GuiObject, options, innerText, event)
    }

    getValue() {
        return this.ctrl.Value
    }
}

class addReactiveCheckBox extends ReactiveControl {
    __New(depend, GuiObject, options, innerText, event := 0) {
        super.__New("CheckBox", depend, GuiObject, options, innerText, event)
    }

    getValue() {
        return this.ctrl.Value
    }

    setValue(newValue) {
        this.ctrl.Value := newValue is Func
            ? newValue(this.ctrl.Value)
            : newValue
    }
}

class addReactiveComboBox extends ReactiveControl {
    __New(depend, GuiObject, options, items, event := 0) {
        this.items := items
        this.vals := []
        this.texts := []
        for val, text in items {
            this.vals.Push(val)
            this.texts.Push(text)
        }
        super.__New("ComboBox", depend, GuiObject, options, this.texts, event)
    }

    getValue() {
        return this.vals[this.ctrl.Value]
    }
}

class addReactiveText extends ReactiveControl {
    __New(depend, GuiObject, options, innerText, event := 0) {
        super.__New("Text", depend, GuiObject, options, innerText, event)
    }
}

class addReactiveRadio extends ReactiveControl {
    __New(depend, GuiObject, options, innerText, event := 0) {
        super.__New("Radio", depend, GuiObject, options, innerText, event)
    }
}