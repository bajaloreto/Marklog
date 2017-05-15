Public Class ComboBoxText
    Inherits ComboBox

    Private UserSelected As Boolean
    Private UserEdited As Boolean

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)
        Me.UserSelected = True
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)


        If e.KeyCode <> Keys.Tab And e.KeyCode <> Keys.Enter Then
            If UserEdited = False Then UndoRedo.UndoBuffer.OldValue = Me.Text
            UserEdited = True
        End If
    End Sub

    Protected Overrides Sub OnSelectedIndexChanged(ByVal e As System.EventArgs)
        MyBase.OnSelectedIndexChanged(e)

        If CurrentControl Is Me And Me.UserSelected = True Then
            UndoRedo.OptionChanged(Me.SelectedItem)

            Me.UserSelected = False
        End If
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)

        If UserEdited = True Then
            UndoRedo.TextChanged(Me.Text)
            UserEdited = False
        End If

        Me.UserSelected = False
    End Sub
End Class
