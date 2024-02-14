class ReactiveControl {
    __New(controlType, GuiObject, options, innerText, event := 0, bind := 0) {
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
            this.ctrl.OnEvent(this.event, (*) => this.callback.Call())
        }
    }

    setOptions(newOptions) {
        this.ctrl.Opt(newOptions)
    }

    getInnerText() {
        return this.ctrl.Text
    }

    setInnerText(newInnerText) {
        this.ctrl.Text := newInnerText
    }

    setEvent(newEvent) {
        this.ctrl.OnEvent(newEvent.event, (*) => newEvent.callback())
    }
}

class ReactiveButton extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Button", GuiObject, options, innerText, event)
    }
}

class ReactiveEdit extends ReactiveControl {
    __New(GuiObject, options, innerText, event := 0) {
        super.__New("Edit", GuiObject, options, innerText, event)
    }
}