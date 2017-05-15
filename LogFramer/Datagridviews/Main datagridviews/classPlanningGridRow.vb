Public Class PlanningGridRow
    Private objGuidReferenceMoment As Guid
    Private objStruct As Struct
    Private objKeyMoment As KeyMoment
    Private strStructSort As String
    Private intStructIndent As Integer
    Private bmStructImage As Bitmap
    Private boolStructImageDirty As Boolean = True
    Private intStructHeight As Integer
    Private datStartDate As Date
    Private datEndDate As Date
    Private datStartPreparation As Date
    Private boolPreparationRepeat As Boolean
    Private datEndFollowUp As Date
    Private boolFollowUpRepeat As Boolean
    Private lstRepeatStartDates As New List(Of Date), lstRepeatEndDates As New List(Of Date)
    Private strOrganisation As String
    Private strLocation As String
    Private intType As Integer
    Private boolIsKeyMoment As Boolean
    Private boolLinksLoaded As Boolean
    Private objOutgoingLinksIndices As New List(Of Integer)
    Private objIncomingLinkIndices As New List(Of Integer)
    Private objKeyMomentLinksAsDate As New List(Of Date)
    Private objActivityLinksAsDate As New List(Of Date)
    Private intRowType As Integer

    Public Enum RowTypes
        Activity = 0
        KeyMoment = 1
        RepeatPurpose = 2
        RepeatOutput = 3
    End Enum

    Public Property GuidReferenceMoment() As Guid
        Get
            Return objGuidReferenceMoment
        End Get
        Set(ByVal value As Guid)
            objGuidReferenceMoment = value
        End Set
    End Property

    Public Property SortNumber() As String
        Get
            Return strStructSort
        End Get
        Set(ByVal value As String)
            strStructSort = value
        End Set
    End Property

    Public Property Indent As Integer
        Get
            Return intStructIndent
        End Get
        Set(value As Integer)
            intStructIndent = value
        End Set
    End Property

    Public Property Struct As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
            If objStruct IsNot Nothing Then objKeyMoment = Nothing
        End Set
    End Property

    Public Property StructImage As Bitmap
        Get
            Return bmStructImage
        End Get
        Set(ByVal value As Bitmap)
            bmStructImage = value
        End Set
    End Property

    Public Property StructImageDirty As Boolean
        Get
            Return boolStructImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolStructImageDirty = value
        End Set
    End Property

    Public Property StructHeight As Integer
        Get
            Return intStructHeight
        End Get
        Set(ByVal value As Integer)
            intStructHeight = value
        End Set
    End Property

    Public Property KeyMoment As KeyMoment
        Get
            Return objKeyMoment
        End Get
        Set(value As KeyMoment)
            objKeyMoment = value
            If objKeyMoment IsNot Nothing Then objStruct = Nothing
        End Set
    End Property

    Public ReadOnly Property IsKeyMoment() As Boolean
        Get
            If KeyMoment Is Nothing Then Return False Else Return True
        End Get
    End Property

    Public ReadOnly Property StartDate As Date
        Get
            Dim selDate As New Date
            If IsKeyMoment = True Then
                Return KeyMoment.ExactDateKeyMoment
            Else
                Dim selActivity As Activity = TryCast(Struct, Activity)
                If selActivity IsNot Nothing Then
                    selDate = selActivity.ActivityDetail.StartDateMainActivity
                End If
            End If
            Return selDate
        End Get
    End Property

    Public ReadOnly Property EndDate As Date
        Get
            Dim selDate As New Date
            If IsKeyMoment = True Then
                Return KeyMoment.ExactDateKeyMoment
            Else
                Dim selActivity As Activity = TryCast(Struct, Activity)
                If selActivity IsNot Nothing Then
                    selDate = selActivity.ActivityDetail.EndDateMainActivity
                End If
            End If
            Return selDate
        End Get
    End Property

    Public Property RowType As Integer
        Get
            Return intRowType
        End Get
        Set(ByVal value As Integer)
            intRowType = value
        End Set
    End Property

    Public Property LinksLoaded As Boolean
        Get
            Return boolLinksLoaded
        End Get
        Set(value As Boolean)
            boolLinksLoaded = value
        End Set
    End Property

    Public Property IncomingLinkIndices As List(Of Integer)
        Get
            Return objIncomingLinkIndices
        End Get
        Set(value As List(Of Integer))
            objIncomingLinkIndices = value
        End Set
    End Property

    Public Property OutgoingLinksIndices As List(Of Integer)
        Get
            Return objOutgoingLinksIndices
        End Get
        Set(value As List(Of Integer))
            objOutgoingLinksIndices = value
        End Set
    End Property

    Public Property KeyMomentLinksAsDate As List(Of Date)
        Get
            Return objKeyMomentLinksAsDate
        End Get
        Set(value As List(Of Date))
            objKeyMomentLinksAsDate = value
        End Set
    End Property

    Public Property ActivityLinksAsDate As List(Of Date)
        Get
            Return objActivityLinksAsDate
        End Get
        Set(value As List(Of Date))
            objActivityLinksAsDate = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal intRowType As Integer)
        Me.RowType = intRowType
    End Sub
End Class

Public Class PlanningGridRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As PlanningGridRow)
        List.Add(gridrow)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As PlanningGridRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(GridRow)
        Else
            List.Insert(index, GridRow)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal gridrow As PlanningGridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As PlanningGridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PlanningGridRow)
            End If
        End Get
        Set(value As PlanningGridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As PlanningGridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As PlanningGridRow) As Boolean
        Return List.Contains(gridrow)
    End Function

    Public Function GetByGuid(ByVal FindGuid As Guid) As PlanningGridRow
        Dim selGridRow As PlanningGridRow
        For Each selGridRow In Me.List
            Select Case selGridRow.RowType
                Case PlanningGridRow.RowTypes.KeyMoment
                    If selGridRow.KeyMoment.Guid = FindGuid Then Return selGridRow
                Case PlanningGridRow.RowTypes.Activity
                    Dim selActivity As Activity = DirectCast(selGridRow.Struct, Activity)
                    If selActivity.Guid = FindGuid Then Return selGridRow
            End Select
        Next
        Return Nothing
    End Function

    Public Function GetReferenceIndex(ByVal selRow As PlanningGridRow) As Integer
        Dim intIndex As Integer = -1

        For Each objRow As PlanningGridRow In Me.List
            If objRow.GuidReferenceMoment = selRow.GuidReferenceMoment Then intIndex += 1
            If objRow Is selRow Then Exit For
        Next

        Return intIndex
    End Function

    Public Function GetPreviousStruct(ByVal intGridRowIndex As Integer) As Struct
        intGridRowIndex -= 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As PlanningGridRow = Me(i)
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Exit For
            End If
        Next

        Return selStruct
    End Function

    Public Function GetNextStruct(ByVal intGridRowIndex As Integer) As Struct
        intGridRowIndex += 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As PlanningGridRow = Me(i)
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Exit For
            End If
        Next

        Return selStruct
    End Function

    Public Function GetPreviousOutput(ByVal intGridRowIndex As Integer) As Output
        intGridRowIndex -= 1
        Dim selOutput As Output = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As PlanningGridRow = Me(i)
            If selGridRow.Struct IsNot Nothing AndAlso selGridRow.Struct.GetType Is GetType(Output) Then
                selOutput = DirectCast(selGridRow.Struct, Output)
                Exit For
            End If
        Next

        Return selOutput
    End Function

    Public Function GetPreviousKeyMomentOfOutput(ByVal intGridRowIndex As Integer) As KeyMoment
        intGridRowIndex -= 1
        Dim selKeyMoment As KeyMoment = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As PlanningGridRow = Me(i)
            If selGridRow.KeyMoment IsNot Nothing Then
                selKeyMoment = selGridRow.KeyMoment
                Exit For
            ElseIf selGridRow.RowType = PlanningGridRow.RowTypes.RepeatOutput Or selGridRow.RowType = PlanningGridRow.RowTypes.RepeatPurpose Then
                Exit For
            End If
        Next

        Return selKeyMoment
    End Function
End Class
