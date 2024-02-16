class ReactiveSignal {
    __New(val) {
        this.val := val
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
    }
}

class ReactiveControl {
    __New(controlType, GuiObject, options, innerText, event := 0) {
        this.GuiObject := GuiObject
        this.options := options
        this.innerTextbel := innerText
        this.ctrlType := controlType
        ; event param should be an object like:
        ; {event: "Click", callback: function}
        this.ctrl := this.GuiObject.Add(this.ctrlType, options, innerText)
        if (event != 0) {
            this.event := event.event
            this.callback := event.callback
            this.ctrl.OnEvent(this.event, (*) => event.callback())
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

    setEvent(newEvent) {
        this.ctrl.OnEvent(newEvent.event, (*) => newEvent.callback())
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