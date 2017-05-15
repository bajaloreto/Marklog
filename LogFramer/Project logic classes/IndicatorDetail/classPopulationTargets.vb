Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class PopulationTarget
    Private objGuid As Guid
    Private intIdPopulationTarget As Integer = -1
    Private objPopulationTargetDeadlineGuid, objParentIndicatorGuid, objReferencePopulationTargetGuid As Guid
    Private boolRelative As Boolean
    Private strPopulationTargetName As String
    Private strFormula As String
    Private dblPercentage As Double = 100

    Public Sub New()

    End Sub

    Public Sub New(ByVal populationtargetdeadlineguid As Guid)
        Me.TargetDeadlineGuid = populationtargetdeadlineguid
    End Sub

    Public Sub New(ByVal populationtargetdeadlineguid As Guid, ByVal percentage As Double)
        Me.TargetDeadlineGuid = populationtargetdeadlineguid
        Me.Percentage = percentage
    End Sub

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property idPopulationTarget As Integer
        Get
            Return intIdPopulationTarget
        End Get
        Set(ByVal value As Integer)
            intIdPopulationTarget = value
        End Set
    End Property

    Public Property TargetDeadlineGuid As Guid
        Get
            Return objPopulationTargetDeadlineGuid
        End Get
        Set(ByVal value As Guid)
            objPopulationTargetDeadlineGuid = value
        End Set
    End Property

    Public Property Percentage As Double
        Get
            Return dblPercentage
        End Get
        Set(ByVal value As Double)
            dblPercentage = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_PopulationTarget
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_PopulationTargets
        End Get
    End Property
End Class

Public Class PopulationTargets
    Inherits System.Collections.CollectionBase

    Public Event PopulationTargetAdded(ByVal sender As Object, ByVal e As PopulationTargetAddedEventArgs)

    Public Sub Add(ByVal populationtarget As PopulationTarget)
        List.Add(populationtarget)
        RaiseEvent PopulationTargetAdded(Me, New PopulationTargetAddedEventArgs(populationtarget))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal populationtarget As PopulationTarget)
        List.Insert(intIndex, populationtarget)
        RaiseEvent PopulationTargetAdded(Me, New PopulationTargetAddedEventArgs(populationtarget))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal populationtarget As PopulationTarget)
        If Me.List.Contains(populationtarget) Then
            Me.List.Remove(populationtarget)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PopulationTarget
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), PopulationTarget)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal populationtarget As PopulationTarget) As Integer
        Return List.IndexOf(populationtarget)
    End Function

    Public Function Contains(ByVal populationtarget As PopulationTarget) As Boolean
        Return List.Contains(populationtarget)
    End Function

    Public Function GetPopulationTargetByGuid(ByVal objGuid As Guid) As PopulationTarget
        Dim selPopulationTarget As PopulationTarget = Nothing

        For Each objPopulationTarget As PopulationTarget In Me.List
            If objPopulationTarget.Guid = objGuid Then
                selPopulationTarget = objPopulationTarget
                Exit For
            End If
        Next
        Return selPopulationTarget
    End Function
End Class

Public Class PopulationTargetAddedEventArgs
    Inherits EventArgs

    Public Property PopulationTarget As PopulationTarget

    Public Sub New(ByVal objPopulationTarget As PopulationTarget)
        MyBase.New()

        Me.PopulationTarget = objPopulationTarget
    End Sub
End Class
