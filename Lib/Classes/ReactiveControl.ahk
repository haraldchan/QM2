class ReactiveButton {
    __New(GuiObject, options, label, callback) {
        this.GuiObject := GuiObject
        this.options := options
        this.label := label
        this.callback := callback

        this.btn := this.GuiObject.AddButton(options, label)
        this.btn.OnEvent("Click", this.callback)
    }

    setOptions(newOptions){
        this.btn.Opt(newOptions)
    }

    setLabel(newLabel){
        this.btn.Text := newLabel
    }

    setCallback(newEvent := "Click", newCallback := this.callback){
        this.btn.OnEvent(newEvent, newCallback)
    }
}

