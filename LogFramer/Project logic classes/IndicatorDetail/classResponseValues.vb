Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ResponseValue

    <ScriptIgnore()> _
    Public WithEvents BooleanValues As New BooleanValues 'one per indicator.responseclass

    Private intIdResponseValue As Integer
    Private objGuid, objParentResponseGuid, objTargetGuid As Guid
    Private strValueName As String
    Private dblValue, dblScoreValue, dblPopulationValue As Double
    Private boolBooleanValue As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByVal targetguid As Guid, ByVal value As Double)
        Me.TargetGuid = targetguid
        Me.Value = value
    End Sub

    Public Property idResponseValue As Integer
        Get
            Return intIdResponseValue
        End Get
        Set(ByVal value As Integer)
            intIdResponseValue = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ParentResponseGuid() As Guid
        Get
            Return objParentResponseGuid
        End Get
        Set(ByVal value As Guid)
            objParentResponseGuid = value
        End Set
    End Property

    Public Property TargetGuid() As Guid
        Get
            Return objTargetGuid
        End Get
        Set(ByVal value As Guid)
            objTargetGuid = value
        End Set
    End Property

    Public Property Value As Double
        Get
            Return dblValue
        End Get
        Set(ByVal value As Double)
            dblValue = value
        End Set
    End Property

    Public Property ScoreValue As Double
        Get
            Return dblScoreValue
        End Get
        Set(ByVal value As Double)
            dblScoreValue = value
        End Set
    End Property

    Public Property PopulationPercentage As Double
        Get
            Return dblPopulationValue
        End Get
        Set(ByVal value As Double)
            dblPopulationValue = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_ResponseValue
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_ResponseValues
        End Get
    End Property

    Private Sub BooleanValues_BooleanValueAdded(ByVal sender As Object, ByVal e As BooleanValueAddedEventArgs) Handles BooleanValues.BooleanValueAdded
        Dim selBooleanValue As BooleanValue = e.BooleanValue

        selBooleanValue.idParent = Me.idResponseValue
        selBooleanValue.ParentGuid = Me.Guid
    End Sub
End Class

Public Class ResponseValues
    Inherits System.Collections.CollectionBase

    Public Event ResponseValueAdded(ByVal sender As Object, ByVal e As ResponseValueAddedEventArgs)

    Public Sub Add(ByVal responsevalue As ResponseValue)
        List.Add(responsevalue)
        RaiseEvent ResponseValueAdded(Me, New ResponseValueAddedEventArgs(responsevalue))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal responsevalue As ResponseValue)
        List.Insert(intIndex, responsevalue)
        RaiseEvent ResponseValueAdded(Me, New ResponseValueAddedEventArgs(responsevalue))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal responsevalue As ResponseValue)
        If Me.List.Contains(responsevalue) Then
            Me.List.Remove(responsevalue)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ResponseValue
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), ResponseValue)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal responsevalue As ResponseValue) As Integer
        Return List.IndexOf(responsevalue)
    End Function

    Public Function Contains(ByVal responsevalue As ResponseValue) As Boolean
        Return List.Contains(responsevalue)
    End Function

    Public Function GetResponseValueByGuid(ByVal objGuid As Guid) As ResponseValue
        Dim selResponseValue As ResponseValue = Nothing

        For Each objResponseValue As ResponseValue In Me.List
            If objResponseValue.Guid = objGuid Then
                selResponseValue = objResponseValue
                Exit For
            End If
        Next
        Return selResponseValue
    End Function
End Class

Public Class ResponseValueAddedEventArgs
    Inherits EventArgs

    Public Property ResponseValue As ResponseValue

    Public Sub New(ByVal objResponseValue As ResponseValue)
        MyBase.New()

        Me.ResponseValue = objResponseValue
    End Sub
End Class


