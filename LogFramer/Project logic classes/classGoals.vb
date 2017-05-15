Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Goal
    Inherits Struct

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Overrides Property Section() As Integer
        Get
            Return 1
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return My.Settings.setStruct1sing
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return My.Settings.setStruct1
        End Get
    End Property
End Class

Public Class Goals
    Inherits Structs

    Public Sub New()

    End Sub

    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Goal
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Goal)
            End If
        End Get
    End Property

    Public Function GetGoalByGuid(ByVal objGuid As Guid) As Goal
        Dim selGoal As Goal = Nothing
        For Each objGoal As Goal In Me.List
            If objGoal.Guid = objGuid Then
                selGoal = objGoal
                Exit For
            End If
        Next
        Return selGoal
    End Function

End Class


