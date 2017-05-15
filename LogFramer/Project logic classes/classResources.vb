Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Resource
    Inherits LogframeObject

    'collections
    <ScriptIgnore()> _
    Public WithEvents BudgetItemReferences As New BudgetItemReferences

    'variables
    Private intIdResource As Integer = -1
    Private intIdStruct As Integer
    Private sngBudgetValue As Single
    Private objParentStructGuid As Guid
    Private bmCellImageBudget As System.Drawing.Bitmap

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

#Region "Properties"
    Public Property idResource As Integer
        Get
            Return intIdResource
        End Get
        Set(ByVal value As Integer)
            intIdResource = value
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

    Public Overrides Property Section() As Integer
        Get
            Return 4
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    Public Property TotalCostAmount() As Single
        Get
            Return sngBudgetValue
        End Get
        Set(ByVal value As Single)
            sngBudgetValue = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Property CellImageBudget() As Bitmap
        Get
            Return bmCellImageBudget
        End Get
        Set(ByVal value As Bitmap)
            bmCellImageBudget = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RowCount() As Integer
        Get
            Dim NmbrRows As Integer
            If Me.BudgetItemReferences.Count > 0 Then
                NmbrRows = Me.BudgetItemReferences.Count + 1
                If My.Settings.setLogframeHideEmptyCells = False Then NmbrRows += 1
            End If
            If NmbrRows = 0 Then NmbrRows = 1
            Return NmbrRows
        End Get
    End Property

    Public Property ParentStructGuid() As Guid
        Get
            Return objParentStructGuid
        End Get
        Set(ByVal value As Guid)
            objParentStructGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Resource
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Resources
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property BudgetName() As String
        Get
            Return LANG_Budget
        End Get
    End Property
#End Region

#Region "Events"
    Private Sub BudgetItemReferences_BudgetItemReferenceAdded(ByVal sender As Object, ByVal e As BudgetItemReferenceAddedEventArgs) Handles BudgetItemReferences.BudgetItemReferenceAdded
        Dim selBudgetItemReference As BudgetItemReference = e.BudgetItemReference

        selBudgetItemReference.idParentResource = Me.idResource
        selBudgetItemReference.ParentResourceGuid = Me.Guid
    End Sub

    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selBudgetItemReference As BudgetItemReference In Me.BudgetItemReferences
            selBudgetItemReference.ParentResourceGuid = Me.Guid
        Next
    End Sub
#End Region

#Region "Methods"
    Public Sub SetTotalCostAmount()
        Me.TotalCostAmount = Me.BudgetItemReferences.GetTotalCost.Amount
    End Sub
#End Region
End Class

Public Class Resources
    Inherits System.Collections.CollectionBase

    Public Event ResourceAdded(ByVal sender As Object, ByVal e As ResourceAddedEventArgs)

#Region "Properties"
    Public ReadOnly Property Section() As Integer
        Get
            Return 4
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Resource
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Resource)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Section As Integer)

    End Sub

    Public Sub Add(ByVal resource As Resource)
        List.Add(resource)
        RaiseEvent ResourceAdded(Me, New ResourceAddedEventArgs(resource))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal resource As Resource)
        If intIndex > Count Or intIndex < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, resource.ItemName))
        ElseIf intIndex = Count Then
            List.Add(resource)
            RaiseEvent ResourceAdded(Me, New ResourceAddedEventArgs(resource))
        Else
            List.Insert(intIndex, resource)
            RaiseEvent ResourceAdded(Me, New ResourceAddedEventArgs(resource))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, Resource.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal resource As Resource)
        If List.Contains(resource) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, resource.ItemName))
        Else
            List.Remove(resource)
        End If
    End Sub

    Public Function IndexOf(ByVal resource As Resource) As Integer
        Return List.IndexOf(resource)
    End Function

    Public Function Contains(ByVal resource As Resource) As Boolean
        Return List.Contains(resource)
    End Function
#End Region

#Region "Get by GUID"
    Public Function GetResourceByGuid(ByVal objGuid As Guid) As Resource
        Dim selResource As Resource = Nothing
        For Each objResource As Resource In Me.List
            If objResource.Guid = objGuid Then
                selResource = objResource
                Exit For
            End If
        Next
        Return selResource
    End Function
#End Region

End Class

Public Class ResourceAddedEventArgs
    Inherits EventArgs

    Public Property Resource As Resource

    Public Sub New(ByVal objResource As Resource)
        MyBase.New()

        Me.Resource = objResource
    End Sub
End Class
