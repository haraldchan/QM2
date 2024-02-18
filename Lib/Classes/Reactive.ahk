; class ReactiveUpdater {
;     static states := []
;     static subEffects := []

;     static sync(signal, newVal){
;         for state in this.states {
;             state[signal] := newVal
;         }
;     }
; }

class ReactiveSignal {
    __New(val) {
        this.val := val
        ; ReactiveUpdater.states.Push(Map(this, this.val))
    }

    get(deriveFunction := 0) {
        if (deriveFunction != 0 && deriveFunction is Func) {
            return deriveFunction(this.val)
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
        
        ; ReactiveUpdater.sync(this, this.val)
    }
}

class ReactiveControl {
    __New(controlType, GuiObject, options, innerText, event := 0) {
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

    getValue() {
        return this.ctrl.Value
    }
    
    setValue(newValue) {
        this.ctrl.Value := newValue is Func
            ? newValue(this.ctrl.Value)
            : newValue
    }

    setEvent(event, callback) {
        this.ctrl.OnEvent(event, (*) => callback())
    }

    disable(boolean){
        this.ctrl.Enabled := !boolean
    }
}

class addReactiveButton extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Button", GuiObject, options, innerText, event)
    }
}

class addReactiveEdit extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Edit", GuiObject, options, innerText, event)
    }
}

class addReactiveCheckBox extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("CheckBox", GuiObject, options, innerText, event)
    }
}

class addReactiveComboBox extends ReactiveControl {
    __New(GuiObject, options, items, event := 0) {
        this.items := items
        this.vals := []
        this.texts := []
        for val, text in items {
            this.vals.Push(val)
            this.texts.Push(text)
        }
        super.__New("ComboBox", GuiObject, options, this.texts, event)
    }

    getValue() {
        return this.vals[this.ctrl.Value]
    }
}

class addReactiveText extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Text", GuiObject, options, innerText, event)
    }
}

class addReactiveRadio extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Radio", GuiObject, options, innerText, event)
    }
}