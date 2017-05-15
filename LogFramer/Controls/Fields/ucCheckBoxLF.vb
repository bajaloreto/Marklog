Public Class CheckBoxLF
    Inherits CheckBox

    Private intType As Integer

    Public Enum Types As Integer
        Normal
        DurationFromStartOfProject
        DurationUntilEndOfProject
    End Enum

    Public Property Type As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property


    Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
        MyBase.OnEnter(e)

        CurrentControl = Me
    End Sub

    Protected Overrides Sub OnCheckedChanged(ByVal e As System.EventArgs)
        MyBase.OnCheckedChanged(e)

        

    End Sub

    Protected Overrides Sub OnValidated(ByVal e As System.EventArgs)
        MyBase.OnValidated(e)

        'CurrentControl = Me

        Select Case Me.Type
            Case Types.Normal
                UndoRedo.OptionChecked(Me.Checked)
            Case Types.DurationFromStartOfProject
                UndoRedo.DurationFromStartOfProject(Me.Checked)
            Case Types.DurationUntilEndOfProject
                UndoRedo.DurationUntilEndOfProject(Me.Checked)
        End Select
    End Sub
End Class
