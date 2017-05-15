Public Class TextBoxLF
    Inherits TextBox

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        If CurrentControl Is Me Then
            UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.TextChanged)
        End If
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)
        UndoRedo.CurrentControlValidated(Me)
    End Sub
End Class
