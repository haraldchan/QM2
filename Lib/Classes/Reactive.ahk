class ReactiveSignal {
    __New(val) {
        this.val := val
        this.subs := []
        return this.val
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
    }

    addSub(controlInstance) {
        this.subs.Push(controlInstance)
    }
}

; class DerivedSignal extends ReactiveSignal {
;     __New(depend, mutation){
;         this.val := depend.get()
;         this.subs := []
;         this.depend := depend
;         this.mutation := mutation
;     }



; }

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
        if (Type(this.ctrl) = "Gui.Text" || Type(this.ctrl) = "Gui.Button" ) {
            this.ctrl.Text := newValue
        } else if (Type(this.ctrl) = "Gui.Edit") {
            this.ctrl.Value := newValue
        }
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