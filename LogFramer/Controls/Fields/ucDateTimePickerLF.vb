Public Class DateTimePickerLF
    Inherits DateTimePicker

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnValueChanged(ByVal eventargs As System.EventArgs)
        MyBase.OnValueChanged(eventargs)

        UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.DateChanged)
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)
        UndoRedo.DateChanged(Me.Value)
    End Sub
End Class
