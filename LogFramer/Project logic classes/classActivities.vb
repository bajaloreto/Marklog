Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Activity
    Inherits Struct

    Private intIdParent As Integer
    Private objParentActivityGuid, objParentOutputGuid As Guid
    Private objActivityDetail As New ActivityDetail

    <ScriptIgnore()> _
    Public WithEvents Resources As New Resources

    <ScriptIgnore()> _
    Public WithEvents Activities As New Activities

#Region "Basic properties"
    Public Property ActivityDetail As ActivityDetail
        Get
            Return objActivityDetail
        End Get
        Set(ByVal value As ActivityDetail)
            objActivityDetail = value
        End Set
    End Property

    Public Overrides Property idStruct As Integer
        Get
            Return MyBase.idStruct
        End Get
        Set(ByVal value As Integer)
            MyBase.idStruct = value

            If value > 0 Then
                Me.ActivityDetail.idStruct = value
            End If
        End Set
    End Property

    Public Overrides Property Section() As Integer
        Get
            Return 4
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactStartDate() As Date
        Get
            If Me.IsProcess = False Then
                Return Me.ActivityDetail.StartDateMainActivity
            Else
                Return Me.GetProcessStartDate
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactEndDate() As Date
        Get
            If Me.IsProcess = False Then
                Return Me.ActivityDetail.EndDateRepeats
            Else
                Return Me.GetProcessEndDate
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property IsProcess As Boolean
        Get
            If Activities.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ActivityType As String
        Get
            If Me.IsProcess Then
                Return LANG_Process
            Else
                Return LANG_Activity
            End If
        End Get
    End Property

    Public Property ParentOutputGuid() As Guid
        Get
            Return objParentOutputGuid
        End Get
        Set(ByVal value As Guid)
            objParentOutputGuid = value
            objParentActivityGuid = Guid.Empty
        End Set
    End Property

    Public Property ParentActivityGuid() As Guid
        Get
            Return objParentActivityGuid
        End Get
        Set(ByVal value As Guid)
            objParentActivityGuid = value
            objParentOutputGuid = Guid.Empty
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return My.Settings.setStruct4sing
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return My.Settings.setStruct4
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Function GetProcessStartDate()
        Dim datStartDate As Date = Date.MaxValue

        If Me.Activities.Count > 0 Then
            For Each selActivity As Activity In Me.Activities
                If selActivity.IsProcess = False Then
                    With selActivity.ActivityDetail
                        If .StartDateMainActivity > Date.MinValue And .StartDateMainActivity < datStartDate Then _
                            datStartDate = .StartDateMainActivity
                    End With

                Else
                    Dim datProcessStartDate As Date = selActivity.GetProcessStartDate
                    If datProcessStartDate > Date.MinValue And datProcessStartDate < datStartDate Then _
                        datStartDate = datProcessStartDate
                End If
            Next
        End If

        If datStartDate = Date.MaxValue Then datStartDate = Date.MinValue

        Return datStartDate
    End Function

    Public Function GetProcessEndDate()
        Dim datEndDate As Date

        If Me.Activities.Count > 0 Then
            For Each selActivity As Activity In Me.Activities
                If selActivity.IsProcess = False Then
                    If selActivity.ActivityDetail.RepeatTimes > 0 And selActivity.ActivityDetail.RepeatEndDates.Count > 0 Then
                        If selActivity.ActivityDetail.EndDateRepeats > datEndDate Then _
                            datEndDate = selActivity.ActivityDetail.EndDateRepeats
                    Else
                        If selActivity.ActivityDetail.EndDateMainActivity > datEndDate Then _
                            datEndDate = selActivity.ActivityDetail.EndDateMainActivity
                    End If
                    
                Else
                    Dim datProcessEndDate As Date = selActivity.GetProcessEndDate
                    If datProcessEndDate > datEndDate Then _
                        datEndDate = datProcessEndDate
                End If
            Next
        End If

        Return datEndDate
    End Function

    Private Sub ActivityDetailToChildActivity()
        Dim NewActivity As New Activity()

        Using copier As New ObjectCopy
            NewActivity = copier.CopyObject(Me)
        End Using

        Me.Indicators.Clear()
        Me.Assumptions.Clear()

        NewActivity.Activities.Clear()

        Me.Activities.Insert(0, NewActivity)
        Me.ActivityDetail = New ActivityDetail
    End Sub
#End Region

#Region "Events"
    Private Sub Activities_ActivityAdded(sender As Object, e As ActivityAddedEventArgs) Handles Activities.ActivityAdded
        Dim selActivity As Activity = e.Activity

        selActivity.idParent = Me.idStruct
        selActivity.ParentActivityGuid = Me.Guid
    End Sub

    'Private Sub Activities_AddedToProcess(ByVal sender As Object, ByVal e As AddedToProcessEventArgs) Handles Activities.AddedToProcess
    '    Dim selActivity As Activity = e.Activity

    '    selActivity.idParent = Me.idStruct
    '    selActivity.ParentActivityGuid = Me.Guid


    '    'if the user already filled out information about the activity, move that information to a new child activity and put it on the first place
    '    If Me.Activities.Count = 1 Then
    '        With Me.ActivityDetail
    '            If .Relative = False Then
    '                If .StartDate <> Date.MinValue Then
    '                    ActivityDetailToChildActivity()
    '                End If
    '            Else
    '                If .PeriodUnit <> ActivityDetail.DurationUnits.NotDefined And .PeriodDirection <> ActivityDetail.DirectionValues.NotDefined Then
    '                    ActivityDetailToChildActivity()
    '                End If
    '                If selActivity.ActivityDetail.GuidReferenceMoment = Me.Guid Then
    '                    selActivity.ActivityDetail.GuidReferenceMoment = Me.Activities(0).Guid
    '                End If
    '            End If
    '        End With
    '    End If
    'End Sub

    Private Sub Resources_ResourceAdded(ByVal sender As Object, ByVal e As ResourceAddedEventArgs) Handles Resources.ResourceAdded
        Dim selResource As Resource = e.Resource

        selResource.idStruct = Me.idStruct
        selResource.ParentStructGuid = Me.Guid
    End Sub
#End Region
End Class

Public Class Activities
    Inherits Structs

    Public Event ActivityAdded(ByVal sender As Object, ByVal e As ActivityAddedEventArgs)
    Public Event AddedToProcess(ByVal sender As Object, ByVal e As AddedToProcessEventArgs)

    Public Sub New()

    End Sub

    Public Shadows Sub Add(ByVal activity As Activity)
        If activity IsNot Nothing Then
            List.Add(activity)
            RaiseEvent ActivityAdded(Me, New ActivityAddedEventArgs(activity))
        End If
    End Sub

    Public Overloads Sub Insert(ByVal index As Integer, ByVal activity As Activity)
        MyBase.Insert(index, activity)
        RaiseEvent ActivityAdded(Me, New ActivityAddedEventArgs(activity))
    End Sub

    Public Sub AddToProcess(ByVal activity As Activity)
        If activity IsNot Nothing Then
            List.Add(activity)
            RaiseEvent AddedToProcess(Me, New AddedToProcessEventArgs(activity))
        End If
    End Sub

    'Public Sub InsertIntoProcess(ByVal index As Integer, ByVal activity As Activity)
    '    MyBase.Insert(index, activity)
    '    RaiseEvent AddedToProcess(Me, New AddedToProcessEventArgs(activity))
    'End Sub

    Public ReadOnly Property LocationsList() As List(Of String)
        Get
            Dim lstLocations As New List(Of String)
            For Each selActivity As Activity In Me.List
                If String.IsNullOrEmpty(selActivity.ActivityDetail.Location) = False Then
                    If lstLocations.Contains(selActivity.ActivityDetail.Location) = False Then _
                        lstLocations.Add(selActivity.ActivityDetail.Location)
                End If

            Next
            Return lstLocations
        End Get
    End Property

    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Activity
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Activity)
            End If
        End Get
    End Property

    Public Function GetActivityByGuid(ByVal objGuid As Guid) As Activity
        Dim selActivity As Activity = Nothing
        For Each objActivity As Activity In Me.List
            If objActivity.Guid = objGuid Then
                selActivity = objActivity
                Exit For
            Else
                If objActivity.IsProcess Then
                    Dim ChildActivity As Activity = objActivity.Activities.GetActivityByGuid(objGuid)
                    If ChildActivity IsNot Nothing Then
                        selActivity = ChildActivity
                        Exit For
                    End If
                End If
            End If
        Next
        Return selActivity
    End Function

    Public Function GetResourceByGuid(ByVal objGuid As Guid) As Resource
        Dim selResource As Resource = Nothing
        For Each selActivity As Activity In Me.List
            selResource = selActivity.Resources.GetResourceByGuid(objGuid)
            If selResource IsNot Nothing Then Exit For
        Next
        Return selResource
    End Function

    Public Overloads Function GetIndicatorByGuid(ByVal objGuid As Guid) As Indicator
        Dim selIndicator As Indicator = Nothing
        For Each selActivity As Activity In Me.List
            selIndicator = selActivity.Indicators.GetIndicatorByGuid(objGuid)
            If selIndicator IsNot Nothing Then
                Exit For
            ElseIf selActivity.Activities.Count > 0 Then
                selIndicator = selActivity.Activities.GetIndicatorByGuid(objGuid)

                If selIndicator IsNot Nothing Then Exit For
            End If
        Next

        Return selIndicator
    End Function

    Public Function GetBudgetItemReferenceByGuid(ByVal objGuid As Guid) As BudgetItemReference
        Dim selBudgetItemReference As BudgetItemReference = Nothing
        For Each selActivity As Activity In Me.List
            For Each selResource As Resource In selActivity.Resources
                selBudgetItemReference = selResource.BudgetItemReferences.GetBudgetItemReferenceByGuid(objGuid)
                If selBudgetItemReference IsNot Nothing Then Return selBudgetItemReference
            Next
        Next
        Return selBudgetItemReference
    End Function
End Class

Public Class ActivityAddedEventArgs
    Inherits EventArgs

    Public Property Activity As Activity

    Public Sub New(ByVal objActivity As Activity)
        MyBase.New()

        Me.Activity = objActivity
    End Sub
End Class

Public Class AddedToProcessEventArgs
    Inherits EventArgs

    Public Property Activity As Activity

    Public Sub New(ByVal objActivity As Activity)
        MyBase.New()

        Me.Activity = objActivity
    End Sub
End Class
