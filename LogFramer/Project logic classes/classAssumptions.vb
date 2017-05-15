Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Assumption
    Inherits LogframeObject

    <ScriptIgnore()> _
    Public WithEvents RiskDetail As New RiskDetail

    <ScriptIgnore()> _
    Public WithEvents AssumptionDetail As New AssumptionDetail

    <ScriptIgnore()> _
    Public WithEvents DependencyDetail As New DependencyDetail

    Private intIdAssumption As Integer = -1
    Private intIdStruct As Integer
    Private intRaidType As Integer
    Private strImpact As String
    Private strResponseStrategy As String
    Private strOwner As String

    Private objParentStructGuid As Guid

#Region "Enumerations"
    Public Enum RaidTypes As Integer
        Risk = 0
        Assumption = 1
        Issue = 2
        Dependency = 3
    End Enum
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Sub New(ByVal Section As Integer, ByVal RTF As String)
        Me.Section = Section
        Me.RTF = RTF
    End Sub
#End Region

#Region "Properties"
    Public Property idAssumption As Integer
        Get
            Return intIdAssumption
        End Get
        Set(ByVal value As Integer)
            intIdAssumption = value
        End Set
    End Property

    Public Property idStruct As Integer
        Get
            Return intIdStruct
        End Get
        Set(ByVal value As Integer)
            intIdStruct = value
        End Set
    End Property

    Public Property ParentStructGuid() As Guid
        Get
            Return objParentStructGuid
        End Get
        Set(ByVal value As Guid)
            objParentStructGuid = value
        End Set
    End Property

    Public Property RaidType As Integer
        Get
            Return intRaidType
        End Get
        Set(value As Integer)
            intRaidType = value
        End Set
    End Property

    Public Property ResponseStrategy() As String
        Get
            Return strResponseStrategy
        End Get
        Set(ByVal value As String)
            strResponseStrategy = value
        End Set
    End Property

    Public Property Owner() As String
        Get
            Return strOwner
        End Get
        Set(ByVal value As String)
            strOwner = value
        End Set
    End Property

    Public Property Impact() As String
        Get
            Return strImpact
        End Get
        Set(ByVal value As String)
            strImpact = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Assumption
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Assumptions
        End Get
    End Property
#End Region

#Region "Events"
    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        'no child collections
    End Sub
#End Region

End Class

Public Class Assumptions
    Inherits System.Collections.CollectionBase

    Private intSection As Integer

    Public Event AssumptionAdded(ByVal sender As Object, ByVal e As AssumptionAddedEventArgs)

#Region "Properties"
    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Assumption
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Assumption)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Section As Integer)
        Me.Section = Section
    End Sub

    Public Sub Add(ByVal assumption As Assumption)
        List.Add(assumption)
        RaiseEvent AssumptionAdded(Me, New AssumptionAddedEventArgs(assumption))
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal assumption As Assumption)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, assumption.ItemName))
        ElseIf index = Count Then
            List.Add(assumption)
            RaiseEvent AssumptionAdded(Me, New AssumptionAddedEventArgs(assumption))
        Else
            List.Insert(index, assumption)
            RaiseEvent AssumptionAdded(Me, New AssumptionAddedEventArgs(assumption))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Assumption.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal assumption As Assumption)
        If List.Contains(assumption) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, assumption.ItemName))
        Else
            List.Remove(assumption)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Public Function IndexOf(ByVal assumption As Assumption) As Integer
        Return List.IndexOf(assumption)
    End Function

    Public Function Contains(ByVal assumption As Assumption) As Boolean
        Return List.Contains(assumption)
    End Function
#End Region

#Region "Get by GUID"
    Public Function GetAssumptionByGuid(ByVal objGuid As Guid) As Assumption
        Dim selAssumption As Assumption = Nothing
        For Each objAssumption As Assumption In Me.List
            If objAssumption.Guid = objGuid Then
                selAssumption = objAssumption
                Exit For
            End If
        Next
        Return selAssumption
    End Function
#End Region

End Class

Public Class AssumptionAddedEventArgs
    Inherits EventArgs

    Public Property Assumption As Assumption

    Public Sub New(ByVal objAssumption As Assumption)
        MyBase.New()

        Me.Assumption = objAssumption
    End Sub
End Class
