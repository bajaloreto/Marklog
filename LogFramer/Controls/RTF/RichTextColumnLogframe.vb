Imports System
Imports System.Windows.Forms

Public Class RichTextCellLogframe
    Inherits DataGridViewTextBoxCell

    Public Sub New()
        'moet niets geïnitialiseerd worden

    End Sub

    Public Overrides Sub InitializeEditingControl(ByVal rowIndex As Integer, ByVal initialFormattedValue As Object, ByVal dataGridViewCellStyle As DataGridViewCellStyle)

        ' Set the value of the editing control to the current cell value.
        MyBase.InitializeEditingControl(rowIndex, initialFormattedValue, dataGridViewCellStyle)

        Dim ctl As RichTextEditingControlLogframe = CType(DataGridView.EditingControl, RichTextEditingControlLogframe)

        Try
            If Me.Value IsNot Nothing Then
                Dim strText As String = CType(Me.Value, String)
                If strText.StartsWith("{\rtf1") Then
                    ctl.Rtf = strText
                Else
                    ctl.Text = strText
                End If
            Else
                ctl.Text = String.Empty
            End If

            'If ctl.Text.Contains("{") Then ctl.Rtf = String.Empty
            If String.IsNullOrEmpty(ctl.Text) Then
                ctl.Font = CurrentText.Font
                ctl.SelectionFont = CurrentText.Font

            End If
            ctl.ScrollBars = RichTextBoxScrollBars.None
            ctl.BeforeFirstStroke = True

        Catch ex As Exception

        End Try

    End Sub

    Public Overrides ReadOnly Property EditType() As Type
        Get
            ' Return the type of the editing contol that RTF Cell uses.
            Return GetType(RichTextEditingControlLogframe)
        End Get
    End Property

    Public Overrides ReadOnly Property ValueType() As Type
        Get
            Return GetType(String)
        End Get
    End Property

    Public Overrides ReadOnly Property DefaultNewRowValue() As Object
        Get
            Return String.Empty
        End Get
    End Property

End Class

Public Class RichTextColumnLogframe
    Inherits DataGridViewColumn

    Public Sub New()
        MyBase.New(New RichTextCellLogframe())

    End Sub

    Public Overrides Property CellTemplate() As DataGridViewCell
        Get
            Return MyBase.CellTemplate
        End Get
        Set(ByVal value As DataGridViewCell)

            If (value IsNot Nothing) AndAlso Not value.GetType().IsAssignableFrom(GetType(RichTextCellLogframe)) Then
                Throw New InvalidCastException("Must be a RichTextCellLogframe")
            End If
            MyBase.CellTemplate = value

        End Set
    End Property

End Class


