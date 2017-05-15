Public Class LogframeRow
    Private intSection As Integer
    Private objStruct As Struct
    Private objIndicator As Indicator
    Private objVerificationSource As VerificationSource
    Private objResource As Resource
    Private objAssumption As Assumption
    Private strStructSort As String
    Private strIndSort As String
    Private strVerSort As String
    Private strRscSort As String
    Private strBudget As String
    Private strAsmSort As String
    Private bmStructImage, bmIndImage As Bitmap
    Private boolStructImageDirty As Boolean = True
    Private boolIndImageDirty As Boolean = True
    Private boolVerImageDirty As Boolean = True
    Private boolRscImageDirty As Boolean = True
    Private boolAsmImageDirty As Boolean = True
    Private intStructHeight, intIndHeight, intVerHeight, intRscHeight, intAsmHeight As Integer
    Private intStructIndent, intIndicatorIndent As Integer
    Private intRowType As Integer

    Public boolTest As Boolean

    Public Enum RowTypes
        Normal = 0
        Section = 1
        RepeatPurpose = 2
        RepeatOutput = 3
    End Enum

#Region "Properties"
    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property Struct As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
        End Set
    End Property

    Public Property Indicator As Indicator
        Get
            Return objIndicator
        End Get
        Set(ByVal value As Indicator)
            objIndicator = value
        End Set
    End Property

    Public Property VerificationSource As VerificationSource
        Get
            Return objVerificationSource
        End Get
        Set(ByVal value As VerificationSource)
            objVerificationSource = value
        End Set
    End Property

    Public Property Resource As Resource
        Get
            Return objResource
        End Get
        Set(ByVal value As Resource)
            objResource = value
        End Set
    End Property

    Public Property Assumption As Assumption
        Get
            Return objAssumption
        End Get
        Set(ByVal value As Assumption)
            objAssumption = value
        End Set
    End Property

    Public Property StructSort As String
        Get
            Return strStructSort
        End Get
        Set(ByVal value As String)
            strStructSort = value
        End Set
    End Property

    Public Property IndicatorSort As String
        Get
            Return strIndSort
        End Get
        Set(ByVal value As String)
            strIndSort = value
        End Set
    End Property

    Public Property VerificationSourceSort As String
        Get
            Return strVerSort
        End Get
        Set(ByVal value As String)
            strVerSort = value
        End Set
    End Property

    Public Property ResourceSort As String
        Get
            Return strRscSort
        End Get
        Set(ByVal value As String)
            strRscSort = value
        End Set
    End Property

    Public Property AssumptionSort As String
        Get
            Return strAsmSort
        End Get
        Set(ByVal value As String)
            strAsmSort = value
        End Set
    End Property

    Public Property StructRtf As String
        Get
            If Me.Struct IsNot Nothing Then
                Return Me.Struct.RTF
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Struct IsNot Nothing Then
                Me.Struct.RTF = value
            End If
        End Set
    End Property

    Public Property IndicatorRtf As String
        Get
            If Me.Indicator IsNot Nothing Then
                Return Me.Indicator.RTF
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Indicator IsNot Nothing Then
                Me.Indicator.RTF = value
            End If
        End Set
    End Property

    Public Property VerificationSourceRtf As String
        Get
            If Me.VerificationSource IsNot Nothing Then
                Return Me.VerificationSource.RTF
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.VerificationSource IsNot Nothing Then
                Me.VerificationSource.RTF = value
            End If
        End Set
    End Property

    Public Property ResourceRtf As String
        Get
            If Me.Resource IsNot Nothing Then
                Return Me.Resource.RTF
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Resource IsNot Nothing Then
                Me.Resource.RTF = value
            End If
        End Set
    End Property

    Public Property TotalCostAmount As Single
        Get
            If Me.Resource IsNot Nothing Then
                Return Me.Resource.TotalCostAmount
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Single)
            If Me.Resource IsNot Nothing Then
                Me.Resource.TotalCostAmount = value
            End If
        End Set
    End Property

    Public Property AssumptionRtf As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.RTF
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Assumption IsNot Nothing Then
                Me.Assumption.RTF = value
            End If
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

    Public Property IndicatorImage As Bitmap
        Get
            Return bmIndImage
        End Get
        Set(ByVal value As Bitmap)
            bmIndImage = value
        End Set
    End Property

    Public Property VerificationSourceImage As Bitmap
        Get
            If VerificationSource IsNot Nothing Then
                Return VerificationSource.CellImage
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Bitmap)
            If VerificationSource IsNot Nothing Then
                VerificationSource.CellImage = value
            End If
        End Set
    End Property

    Public Property ResourceImage As Bitmap
        Get
            If Resource IsNot Nothing Then
                Return Resource.CellImage
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Bitmap)
            If Resource IsNot Nothing Then
                Resource.CellImage = value
            End If
        End Set
    End Property

    Public Property AssumptionImage As Bitmap
        Get
            If Assumption IsNot Nothing Then
                Return Assumption.CellImage
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Bitmap)
            If Assumption IsNot Nothing Then
                Assumption.CellImage = value
            End If
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

    Public Property IndicatorImageDirty As Boolean
        Get
            Return boolIndImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolIndImageDirty = value
        End Set
    End Property

    Public Property VerificationSourceImageDirty As Boolean
        Get
            Return boolVerImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolVerImageDirty = value
        End Set
    End Property

    Public Property ResourceImageDirty As Boolean
        Get
            Return boolRscImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolRscImageDirty = value
        End Set
    End Property

    Public Property AssumptionImageDirty As Boolean
        Get
            Return boolAsmImageDirty
        End Get
        Set(ByVal value As Boolean)
            boolAsmImageDirty = value
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

    Public Property IndicatorHeight As Integer
        Get
            Return intIndHeight
        End Get
        Set(ByVal value As Integer)
            intIndHeight = value
        End Set
    End Property

    Public Property VerificationSourceHeight As Integer
        Get
            Return intVerHeight
        End Get
        Set(ByVal value As Integer)
            intVerHeight = value
        End Set
    End Property

    Public Property ResourceHeight As Integer
        Get
            Return intRscHeight
        End Get
        Set(ByVal value As Integer)
            intRscHeight = value
        End Set
    End Property

    Public Property AssumptionHeight As Integer
        Get
            Return intAsmHeight
        End Get
        Set(ByVal value As Integer)
            intAsmHeight = value
        End Set
    End Property

    Public Property StructIndent As Integer
        Get
            Return intStructIndent
        End Get
        Set(value As Integer)
            intStructIndent = value
        End Set
    End Property

    Public Property IndicatorIndent As Integer
        Get
            Return intIndicatorIndent
        End Get
        Set(value As Integer)
            intIndicatorIndent = value
        End Set
    End Property

    Public Property RowType As Integer
        Get
            Return intRowType
        End Get
        Set(ByVal value As Integer)
            intRowType = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal intSection As Integer)
        Me.Section = intSection
    End Sub

    Public Sub New(ByVal intSection As Integer, ByVal objStruct As Struct, ByVal objIndicator As Indicator, ByVal objVerificationSource As VerificationSource, ByVal objAssumption As Assumption)
        Me.Section = intSection
        Me.Struct = objStruct
        Me.Indicator = objIndicator
        Me.VerificationSource = objVerificationSource
        Me.Assumption = objAssumption
    End Sub

    Public Sub New(ByVal objGoal As Goal, ByVal objIndicator As Indicator, ByVal objVerificationSource As VerificationSource, ByVal objAssumption As Assumption)
        Me.Section = LogFrame.SectionTypes.GoalsSection
        Me.Struct = objGoal
        Me.Indicator = objIndicator
        Me.VerificationSource = objVerificationSource
        Me.Assumption = objAssumption
    End Sub

    Public Sub New(ByVal objPurpose As Purpose, ByVal objIndicator As Indicator, ByVal objVerificationSource As VerificationSource, ByVal objAssumption As Assumption)
        Me.Section = LogFrame.SectionTypes.PurposesSection
        Me.Struct = objPurpose
        Me.Indicator = objIndicator
        Me.VerificationSource = objVerificationSource
        Me.Assumption = objAssumption
    End Sub

    Public Sub New(ByVal objOutput As Output, ByVal objIndicator As Indicator, ByVal objVerificationSource As VerificationSource, ByVal objAssumption As Assumption)
        Me.Section = LogFrame.SectionTypes.OutputsSection
        Me.Struct = objOutput
        Me.Indicator = objIndicator
        Me.VerificationSource = objVerificationSource
        Me.Assumption = objAssumption
    End Sub

    Public Sub New(ByVal objActivity As Activity, ByVal objIndicator As Indicator, ByVal objVerificationSource As VerificationSource, ByVal objAssumption As Assumption)
        Me.Section = LogFrame.SectionTypes.ActivitiesSection
        Me.Struct = objActivity
        Me.Indicator = objIndicator
        Me.VerificationSource = objVerificationSource
        Me.Assumption = objAssumption
    End Sub

    Public Sub New(ByVal objActivity As Activity, ByVal objResource As Resource, ByVal objAssumption As Assumption)
        Me.Section = LogFrame.SectionTypes.ActivitiesSection
        Me.Struct = objActivity
        Me.Resource = objResource
        Me.Assumption = objAssumption
    End Sub
#End Region
End Class

Public Class LogframeRows
    Inherits System.Collections.CollectionBase

#Region "Properties"
    Default Public Property Item(ByVal index As Integer) As LogframeRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), LogframeRow)
            End If
        End Get
        Set(ByVal value As LogframeRow)
            List.Item(index) = value
        End Set
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal logframerow As LogframeRow)
        If logframerow IsNot Nothing Then
            List.Add(logframerow)
        End If
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal logframerow As LogframeRow)
        If logframerow IsNot Nothing Then
            If index > Count Or index < 0 Then
                System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
            Else
                List.Insert(index, logframerow)
            End If
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal logframerow As LogframeRow)
        If List.Contains(logframerow) = False Then
            System.Windows.Forms.MessageBox.Show("Logframe row not in list")
        Else
            List.Remove(logframerow)
        End If
    End Sub

    Public Sub RemoveStruct(ByVal struct As Struct)
        If struct IsNot Nothing Then
            CurrentLogFrame.RemoveStruct(struct)
        End If
    End Sub

    Public Function IndexOf(ByVal logframerow As LogframeRow) As Integer
        Return List.IndexOf(logframerow)
    End Function

    Public Function Contains(ByVal logframerow As LogframeRow) As Boolean
        Return List.Contains(logframerow)
    End Function
#End Region

#Region "Get"
    Public Function GetRowIndexOfStruct(ByVal SearchStruct As Struct) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Struct Is SearchStruct Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfStruct(ByVal objGuid As Guid) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Struct.Guid = objGuid Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfFirstRowInSection(ByVal intSearchSection As Integer) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.Section = intSearchSection And selLogframeRow.RowType = LogframeRow.RowTypes.Normal Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfSectionTitle(ByVal intSearchSection As Integer) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.Section = intSearchSection And selLogframeRow.RowType = LogframeRow.RowTypes.Section Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfFirstStructInSection(ByVal SearchStruct As Struct, ByVal intSearchSection As Integer) As Integer
        Dim index As Integer = -1
        If SearchStruct IsNot Nothing Then
            Dim lstSection As List(Of LogframeRow) = GetSection(intSearchSection)
            Dim selRow As LogframeRow = Nothing

            For i = 0 To lstSection.Count - 1
                selRow = CType(lstSection(i), LogframeRow)
                If selRow.Struct Is SearchStruct Then
                    index = List.IndexOf(selRow)
                    Exit For
                End If
            Next
        End If
        Return index
    End Function

    Public Function GetRowIndexOfIndicator(ByVal SearchIndicator As Indicator) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Indicator Is SearchIndicator Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfIndicator(ByVal objGuid As Guid) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Indicator.Guid = objGuid Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfResource(ByVal SearchResource As Resource) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Resource Is SearchResource Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfResource(ByVal objGuid As Guid) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Resource.Guid = objGuid Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfVerificationSource(ByVal SearchVerificationSource As VerificationSource) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.VerificationSource Is SearchVerificationSource Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfVerificationSource(ByVal objGuid As Guid) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.VerificationSource.Guid = objGuid Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfAssumption(ByVal SearchAssumption As Assumption) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Assumption Is SearchAssumption Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetRowIndexOfAssumption(ByVal objGuid As Guid) As Integer
        For Each selLogframeRow As LogframeRow In Me.List
            If selLogframeRow.RowType = LogframeRow.RowTypes.Normal And selLogframeRow.Assumption.Guid = objGuid Then
                Return Me.List.IndexOf(selLogframeRow)
            End If
        Next

        Return -1
    End Function

    Public Function GetSection(ByVal SearchSection As Integer) As List(Of LogframeRow)
        Dim lstFound As New List(Of LogframeRow)
        For i = 0 To List.Count - 1
            Dim selLogframeRow As LogframeRow = CType(Me(i), LogframeRow)
            If selLogframeRow.Section = SearchSection And selLogframeRow.RowType <> LogframeRow.RowTypes.Section Then lstFound.Add(selLogframeRow)
        Next
        Return lstFound
    End Function

    Private Function GetRepeatedPurposes() As List(Of LogframeRow)
        Dim lstSectionActivities As List(Of LogframeRow) = GetSection(LogFrame.SectionTypes.ActivitiesSection)
        Dim lstFound As New List(Of LogframeRow)
        For i = 0 To lstSectionActivities.Count - 1
            Dim selLogframeRow As LogframeRow = CType(lstSectionActivities(i), LogframeRow)
            If selLogframeRow.RowType = LogframeRow.RowTypes.RepeatPurpose Then lstFound.Add(selLogframeRow)
        Next
        Return lstFound
    End Function

    Public Function GetPreviousStruct(ByVal intGridRowIndex As Integer, Optional ByVal SearchSection As Integer = -1) As Struct
        intGridRowIndex -= 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)

            If SearchSection >= 0 Then
                If selGridRow.RowType <> LogframeRow.RowTypes.Normal Then Return Nothing
                If selGridRow.Section <> SearchSection Then Return Nothing
            End If
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Return selStruct
            End If
        Next

        Return Nothing
    End Function

    Public Function GetNextStruct(ByVal intGridRowIndex As Integer) As Struct
        intGridRowIndex += 1
        Dim selStruct As Struct = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.Struct IsNot Nothing Then
                selStruct = selGridRow.Struct
                Exit For
            End If
        Next

        Return selStruct
    End Function

    Public Function GetPreviousIndicator(ByVal intGridRowIndex As Integer) As Indicator
        intGridRowIndex -= 1
        Dim selIndicator As Indicator = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.Indicator IsNot Nothing Then
                selIndicator = selGridRow.Indicator
                Exit For
            End If
        Next

        Return selIndicator
    End Function

    Public Function GetPreviousIndicator(ByVal intGridRowIndex As Integer, ByVal SearchSection As Integer) As Indicator
        intGridRowIndex -= 1
        Dim selIndicator As Indicator = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)

            If selGridRow.RowType <> LogframeRow.RowTypes.Normal Then Return Nothing
            If selGridRow.Section <> SearchSection Then Return Nothing
            If selGridRow.Struct IsNot Nothing And selGridRow.Indicator Is Nothing Then Return Nothing

            If selGridRow.Indicator IsNot Nothing Then
                selIndicator = selGridRow.Indicator
                Return selIndicator
            End If
        Next

        Return Nothing
    End Function

    Public Function GetNextIndicator(ByVal intGridRowIndex As Integer) As Indicator
        intGridRowIndex += 1
        Dim selIndicator As Indicator = Nothing

        For i = intGridRowIndex To Me.Count - 1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.Indicator IsNot Nothing Then
                selIndicator = selGridRow.Indicator
                Exit For
            End If
        Next

        Return selIndicator
    End Function

    Public Function GetPreviousResource(ByVal intGridRowIndex As Integer) As Resource
        intGridRowIndex -= 1
        Dim selResource As Resource = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.Resource IsNot Nothing Then
                selResource = selGridRow.Resource
                Exit For
            End If
        Next

        Return selResource
    End Function

    Public Function GetIndexOfPreviousResourceOfStruct(ByVal intGridRowIndex As Integer) As Integer
        
    End Function

    Public Function GetPreviousVerificationSource(ByVal intGridRowIndex As Integer) As VerificationSource
        intGridRowIndex -= 1
        Dim selVerificationSource As VerificationSource = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.VerificationSource IsNot Nothing Then
                selVerificationSource = selGridRow.VerificationSource
                Exit For
            End If
        Next

        Return selVerificationSource
    End Function

    Public Function GetPreviousAssumption(ByVal intGridRowIndex As Integer) As Assumption
        intGridRowIndex -= 1
        Dim selAssumption As Assumption = Nothing

        For i = intGridRowIndex To 0 Step -1
            Dim selGridRow As LogframeRow = Me(i)
            If selGridRow.Assumption IsNot Nothing Then
                selAssumption = selGridRow.Assumption
                Exit For
            End If
        Next

        Return selAssumption
    End Function
#End Region
End Class
