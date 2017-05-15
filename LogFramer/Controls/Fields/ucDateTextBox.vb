Public Class DateTextBox
    Inherits TextBox

    Private ModifyDate As Boolean

    Public Property DateValue() As Date
        Get
            Dim datValue As Date
            If Date.TryParse(Me.Text, datValue) = True Then
                Return datValue
            Else
                Return Date.MinValue
            End If
        End Get
        Set(ByVal value As Date)
            If value > Date.MinValue And value < Date.MaxValue Then
                Me.Text = value.ToString("d")
            Else
                Me.Text = String.Empty
            End If
        End Set
    End Property

    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        ForeColor = Color.Black
        ModifyDate = True
        If Me.DateValue = Date.MinValue Then Me.Text = Now.Date

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnKeyDown(e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)
    End Sub

    Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
        MyBase.OnLeave(e)

        ModifyDate = False
    End Sub

    Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
        MyBase.OnTextChanged(e)
        If ModifyDate = False Then
            Try

                If Me.DateValue = Date.MinValue And Me.ReadOnly = False Then
                    ForeColor = BackColor
                    Visible = True
                ElseIf Me.DateValue = Date.MinValue And Me.ReadOnly = True Then
                    Visible = False
                Else
                    ForeColor = Color.Black
                    Visible = True
                End If
            Catch ex As InvalidCastException
                ForeColor = Color.Black
            End Try
        End If
        If CurrentControl Is Me Then UndoRedo.UndoBuffer_SetAction(classUndoRedo.Actions.DateChanged)
    End Sub

    Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
        MyBase.OnValidating(e)
        If String.IsNullOrEmpty(Me.Text) Then Me.Text = Date.MinValue
    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)
        UndoRedo.CurrentControlValidated(Me)
    End Sub
End Class

