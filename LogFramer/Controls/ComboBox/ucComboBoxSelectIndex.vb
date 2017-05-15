Public Class ComboBoxSelectIndex
    Inherits ComboBox

    Private UserSelected As Boolean

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        MyBase.OnMouseDown(e)
        Me.UserSelected = True
    End Sub

    Protected Overrides Sub OnSelectedIndexChanged(ByVal e As System.EventArgs)
        MyBase.OnSelectedIndexChanged(e)

        If CurrentControl Is Me And Me.UserSelected = True Then
            UndoRedo.OptionChanged(Me.SelectedIndex)
            Me.UserSelected = False
        End If
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)

        Me.UserSelected = False
    End Sub
End Class
