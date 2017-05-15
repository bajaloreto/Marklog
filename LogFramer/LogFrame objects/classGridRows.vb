Public Class GridRow
    Private objParent As GridRows
    Private intSection As Integer
    Private intIndicatorIndex As Integer = -1
    Private intVerificationSourceIndex As Integer = -1
    Private intResourceIndex As Integer = -1
    Private intBudgetItemIndex As Integer = -1
    Private intAssumptionIndex As Integer = -1
    Private classStruct As Struct
    Private typeRowValue As String
    Private intNewlyAdded As Integer

    Private newSIVValue As String
    Private newAsmValue As String
    Private structHeightValue As Integer
    Private indHeightValue As Integer
    Private verHeightValue As Integer
    Private asmHeightValue As Integer
    Private structMergeValue As Integer
    Private indMergeValue As Integer
    Private writeToRowValue As Boolean

    Public Enum RowTypes
        Normal = 0
        Section = 1
        RepeatPurpose = 2
        RepeatOutput = 3
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer)
        Me.Section = section
    End Sub

    Public Sub New(ByVal section As Integer, ByVal struct As Struct)
        Me.Section = section
        Me.Struct = struct
    End Sub

    Public Sub New(ByVal section As Integer, ByVal indicatorindex As Integer, ByVal verificationsourceindex As Integer, _
                   ByVal assumptionindex As Integer, ByVal struct As Struct)
        Me.Section = section
        Me.Struct = struct

        Me.IndicatorIndex = indicatorindex
        Me.VerificationSourceIndex = verificationsourceindex
        Me.AssumptionIndex = assumptionindex
    End Sub

    Public Sub New(ByVal section As Integer, ByVal resourceindex As Integer, ByVal assumptionindex As Integer, ByVal struct As Struct)
        Me.Section = section
        Me.Struct = struct

        Me.ResourceIndex = resourceindex
        Me.AssumptionIndex = assumptionindex
    End Sub

    Public Sub New(ByVal section As Integer, ByVal indicatorindex As Integer, ByVal verificationsourceindex As Integer, _
                   ByVal resourceindex As Integer, ByVal assumptionindex As Integer, ByVal struct As Struct)
        Me.Section = section
        Me.Struct = struct
        Me.IndicatorIndex = indicatorindex
        Me.VerificationSourceIndex = verificationsourceindex
        Me.ResourceIndex = resourceindex
        Me.AssumptionIndex = assumptionindex
    End Sub

#Region "Properties and indexes"
    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property IndicatorIndex() As Integer
        Get
            Return intIndicatorIndex
        End Get
        Set(ByVal value As Integer)
            intIndicatorIndex = value
        End Set
    End Property

    Public Property VerificationSourceIndex() As Integer
        Get
            Return intVerificationSourceIndex
        End Get
        Set(ByVal value As Integer)
            intVerificationSourceIndex = value
        End Set
    End Property

    Public Property ResourceIndex() As Integer
        Get
            Return intResourceIndex
        End Get
        Set(ByVal value As Integer)
            intResourceIndex = value
        End Set
    End Property

    Public Property AssumptionIndex() As Integer
        Get
            Return intAssumptionIndex
        End Get
        Set(ByVal value As Integer)
            intAssumptionIndex = value
        End Set
    End Property

    Public Property RowType() As Integer
        Get
            Return typeRowValue
        End Get
        Set(ByVal value As Integer)
            typeRowValue = value
        End Set
    End Property

    Public Property NewlyAdded() As Integer
        Get
            Return intNewlyAdded
        End Get
        Set(ByVal value As Integer)
            intNewlyAdded = value
        End Set
    End Property

    Public ReadOnly Property MeansBudget() As Boolean
        Get
            If Me.Section = LogFrame.SectionTypes.ActivitiesSection And CurrentProjectForm.dgvLogframe.ShowResourcesBudget = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
#End Region 'Properties and indexes

#Region "Struct columns"
    Public Property Struct() As Struct
        Get
            Return classStruct
        End Get
        Set(ByVal value As Struct)
            classStruct = value
        End Set
    End Property

    Public ReadOnly Property StructSort() As String
        Get
            Dim strSort As String = Nothing
            If Me.Struct IsNot Nothing Then
                If Me.StructRTF <> "" Then _
                    strSort = CurrentLogFrame.GetStructSortNumber(Me.Struct)
            End If
            Return strSort
        End Get
    End Property

    Public Property StructRTF() As String
        Get
            Dim strRTF As String = Nothing

            If Me.Struct IsNot Nothing Then
                If Me.MeansBudget = False Then
                    If Me.IndicatorIndex <= 0 And Me.VerificationSourceIndex <= 0 And Me.AssumptionIndex <= 0 Then
                        strRTF = Me.Struct.RTF
                        If strRTF = "" Then strRTF = StructEmptyRTF()
                    End If
                Else
                    If Me.ResourceIndex <= 0 And Me.AssumptionIndex <= 0 Then
                        strRTF = Me.Struct.RTF
                        If strRTF = "" Then strRTF = StructEmptyRTF()
                    End If
                End If
            End If
            Return strRTF
        End Get
        Set(ByVal value As String)
            If Me.Struct Is Nothing Then
                NewStruct()
            Else
                If Me.Struct.RTF = Nothing Then
                    Select Case intSection
                        Case LogFrame.SectionTypes.PurposesSection
                            NewlyAdded = LogFrame.ObjectTypes.PurposeHidden
                        Case LogFrame.SectionTypes.OutputsSection
                            NewlyAdded = LogFrame.ObjectTypes.OutputHidden
                    End Select
                End If
            End If

            Me.Struct.RTF = value
        End Set
    End Property

    Private Sub NewStruct()
        Select Case intSection
            Case LogFrame.SectionTypes.GoalsSection
                Me.Struct = New Goal()
                NewlyAdded = LogFrame.ObjectTypes.Goal
            Case LogFrame.SectionTypes.PurposesSection
                Me.Struct = New Purpose()
                NewlyAdded = LogFrame.ObjectTypes.Purpose
            Case LogFrame.SectionTypes.OutputsSection
                Me.Struct = New Output()
                NewlyAdded = LogFrame.ObjectTypes.Output
            Case LogFrame.SectionTypes.ActivitiesSection
                Me.Struct = New Activity()
                NewlyAdded = LogFrame.ObjectTypes.Activity
        End Select
    End Sub

    Private Function StructEmptyRTF() As String
        Dim strRTF As String = ""
        Dim strStructSortNumber As String = CurrentLogFrame.GetStructSortNumber(Me.Struct)
        Dim strBegin As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Verdana;}}" & _
                        "{\colortbl ;\red196\green196\blue196;}\viewkind4\uc1\pard\cf1\f0\fs24\{"
        Dim strEnd As String = "\}\cf0\fs23\par}"

        Select Case intSection
            Case LogFrame.SectionTypes.GoalsSection
                strRTF = strBegin & My.Settings.setStruct1sing & " " & strStructSortNumber & strEnd
            Case LogFrame.SectionTypes.PurposesSection
                If Me.Indicator IsNot Nothing Or Me.Assumption IsNot Nothing Then
                    strRTF = strBegin & My.Settings.setStruct2sing & " " & strStructSortNumber & strEnd
                End If
            Case LogFrame.SectionTypes.OutputsSection
                If Me.RowType = GridRow.RowTypes.RepeatPurpose Then
                    strRTF = strBegin & My.Settings.setStruct2sing & " " & strStructSortNumber & strEnd
                Else
                    If Me.Indicator IsNot Nothing Or Me.Assumption IsNot Nothing Then
                        strRTF = strBegin & My.Settings.setStruct3sing & " " & strStructSortNumber & strEnd
                    End If
                End If
            Case LogFrame.SectionTypes.ActivitiesSection
                If Me.RowType = GridRow.RowTypes.RepeatPurpose Then
                    strRTF = strBegin & My.Settings.setStruct2sing & " " & strStructSortNumber & strEnd
                ElseIf Me.RowType = GridRow.RowTypes.RepeatOutput Then
                    strRTF = strBegin & My.Settings.setStruct3sing & " " & strStructSortNumber & strEnd
                Else
                    strRTF = strBegin & My.Settings.setStruct4sing & " " & strStructSortNumber & strEnd
                End If
        End Select

        Return strRTF
    End Function
#End Region 'Struct columns

#Region "Indicator columns"
    Public Property Indicator() As Indicator
        Get
            Dim selInd As Indicator = Nothing
            If Me.Struct IsNot Nothing Then
                If Me.Struct.Indicators.Count > 0 And IndicatorIndex >= 0 Then
                    selInd = Me.Struct.Indicators(intIndicatorIndex)
                Else
                    selInd = Nothing
                End If
            End If
            Return selInd
        End Get
        Set(ByVal value As Indicator)
            If Me.Struct.Indicators(Me.IndicatorIndex) Is Nothing Then
                Me.Struct.Indicators.Add(value)
            Else
                Dim selInd As Indicator = Me.Struct.Indicators(Me.IndicatorIndex)
                selInd = value
            End If
        End Set
    End Property

    Public ReadOnly Property IndSort() As String
        Get
            Dim strSort As String = Nothing
            If Me.Indicator IsNot Nothing Then
                If Me.IndRTF <> "" Then strSort = IndSortNumber()
            End If

            Return strSort
        End Get
    End Property

    Private Function IndSortNumber() As String
        Dim strSort As String = Nothing
        Dim intNr As Integer = Me.Struct.Indicators.IndexOf(Me.Indicator) + 1
        strSort = CurrentLogFrame.GetStructSortNumber(Me.Struct) & "." & intNr.ToString
        Return strSort
    End Function

    Public Property IndRTF() As String
        Get
            Dim strRTF As String = Nothing
            If Me.Indicator IsNot Nothing Then
                If Me.VerificationSourceIndex <= 0 Then
                    strRTF = Me.Indicator.RTF
                    If strRTF = "" Then strRTF = IndEmptyRTF()
                End If
            End If

            Return strRTF
        End Get
        Set(ByVal value As String)
            If Me.Indicator Is Nothing Then NewIndicator()
            Me.Indicator.RTF = value
        End Set
    End Property

    Private Sub NewIndicator()
        If Me.Struct Is Nothing Then NewStruct()
        Me.Struct.Indicators.Add(New Indicator())
        Me.IndicatorIndex = Me.Struct.Indicators.Count - 1
        NewlyAdded = LogFrame.ObjectTypes.Indicator
    End Sub

    Private Function IndEmptyRTF() As String
        Dim strRTF As String
        Dim strBegin As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Verdana;}}" & _
                        "{\colortbl ;\red196\green196\blue196;}\viewkind4\uc1\pard\cf1\f0\fs24\{"
        Dim strEnd As String = "\}\cf0\fs23\par}"
        Dim intNr As Integer = Me.IndicatorIndex + 1

        strRTF = strBegin & Indicator.ItemName & intNr.ToString & strEnd

        Return strRTF
    End Function
#End Region 'Indicator columns

#Region "Verification source columns"
    Public Property VerificationSource() As VerificationSource
        Get
            Dim selVer As VerificationSource = Nothing
            If Me.Indicator IsNot Nothing Then
                If Me.Indicator.VerificationSources.Count > 0 And VerificationSourceIndex >= 0 Then _
                    selVer = Me.Indicator.VerificationSources(intVerificationSourceIndex)
            End If
            Return selVer
        End Get
        Set(ByVal value As VerificationSource)
            If Me.Indicator.VerificationSources(intVerificationSourceIndex) Is Nothing Then
                Me.Indicator.VerificationSources.Add(value)
            Else
                Dim selVer As VerificationSource = Me.Indicator.VerificationSources(intVerificationSourceIndex)
                selVer = value
            End If
        End Set
    End Property

    Public ReadOnly Property VerSort() As String
        Get
            Dim strSort As String = Nothing
            If Me.VerificationSource IsNot Nothing Then
                Dim intNr As Integer = Me.Indicator.VerificationSources.IndexOf(Me.VerificationSource) + 1
                strSort = IndSortNumber() & "." & intNr.ToString
            End If
            Return strSort
        End Get
    End Property

    Public Property VerRTF() As String
        Get
            Dim strRTF As String = Nothing
            If Me.VerificationSource IsNot Nothing Then strRTF = Me.VerificationSource.RTF
            Return strRTF
        End Get
        Set(ByVal value As String)
            If Me.VerificationSource Is Nothing Then NewVerificationSource()
            Me.VerificationSource.RTF = value
        End Set
    End Property

    Private Sub NewVerificationSource()
        If Me.Struct Is Nothing Then NewStruct()
        If Me.Indicator Is Nothing Then NewIndicator()
        Me.Indicator.VerificationSources.Add(New VerificationSource())
        Me.VerificationSourceIndex = Me.Indicator.VerificationSources.Count - 1
        NewlyAdded = LogFrame.ObjectTypes.VerificationSource
    End Sub
#End Region 'Verification source columns

#Region "Resource columns"
    Public Property Resource() As Resource
        Get
            Dim selRsc As Resource = Nothing
            Dim selActivity As Activity = TryCast(Me.Struct, Activity)
            If selActivity IsNot Nothing Then
                If selActivity.Resources.Count > 0 And ResourceIndex >= 0 Then
                    selRsc = selActivity.Resources(intResourceIndex)
                Else
                    selRsc = Nothing
                End If
            End If
            Return selRsc
        End Get
        Set(ByVal value As Resource)
            Dim selActivity As Activity = TryCast(Me.Struct, Activity)
            If selActivity IsNot Nothing Then
                If selActivity.Resources(Me.ResourceIndex) Is Nothing Then
                    selActivity.Resources.Add(value)
                Else
                    Dim selRsc As Resource = selActivity.Resources(Me.ResourceIndex)
                    selRsc = value
                End If
            End If
        End Set
    End Property

    Public ReadOnly Property RscSort() As String
        Get
            Dim strSort As String = Nothing
            If Me.Resource IsNot Nothing Then
                If String.IsNullOrEmpty(Me.RscRTF) = False Then strSort = RscSortNumber()
            End If

            Return strSort
        End Get
    End Property

    Private Function RscSortNumber() As String
        Dim strSort As String = Nothing
        Dim selActivity As Activity = TryCast(Me.Struct, Activity)
        If selActivity IsNot Nothing Then
            Dim intNr As Integer = selActivity.Resources.IndexOf(Me.Resource) + 1
            strSort = String.Format("{0}.{1}", CurrentLogFrame.GetStructSortNumber(Me.Struct), intNr)
        End If

        Return strSort
    End Function

    Public Property RscRTF() As String
        Get
            Dim strRTF As String = Nothing
            If Me.Resource IsNot Nothing Then
                strRTF = Me.Resource.RTF
                If String.IsNullOrEmpty(strRTF) Then strRTF = RscEmptyRTF()
            End If
            Return strRTF
        End Get
        Set(ByVal value As String)
            If Me.Resource Is Nothing Then NewResource()
            Me.Resource.RTF = value
        End Set
    End Property

    Private Sub NewResource()
        If Me.Struct Is Nothing Then NewStruct()
        Dim selActivity As Activity = TryCast(Me.Struct, Activity)

        If selActivity IsNot Nothing Then
            selActivity.Resources.Add(New Resource())
            Me.ResourceIndex = selActivity.Resources.Count - 1
            NewlyAdded = LogFrame.ObjectTypes.Resource
        End If
    End Sub

    Private Function RscEmptyRTF() As String
        Dim strRTF As String
        Dim strBegin As String = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 Verdana;}}" & _
                        "{\colortbl ;\red196\green196\blue196;}\viewkind4\uc1\pard\cf1\f0\fs24\{"
        Dim strEnd As String = "\}\cf0\fs23\par}"
        Dim intNr As Integer = Me.ResourceIndex + 1

        strRTF = strBegin & Resource.ItemName & intNr.ToString & strEnd

        Return strRTF
    End Function
#End Region 'Resource columns

#Region "Budget columns"
    Public Property RscBudget() As Single
        Get
            If Me.Resource IsNot Nothing Then
                Return Me.Resource.TotalCostAmount
            Else
                Return Nothing
            End If
        End Get
        Set(ByVal value As Single)
            If Me.Resource Is Nothing Then NewResource()
            Me.Resource.TotalCostAmount = value
        End Set
    End Property
#End Region 'Budget item columns

#Region "Assumption columns"
    Public Property Assumption() As Assumption
        Get
            Dim selAsm As Assumption = Nothing
            If Me.Struct IsNot Nothing Then
                If Me.Struct.Assumptions.Count > 0 And AssumptionIndex >= 0 Then
                    selAsm = Me.Struct.Assumptions(intAssumptionIndex)
                Else
                    selAsm = Nothing
                End If
            End If
            Return selAsm
        End Get
        Set(ByVal value As Assumption)
            If Me.Struct.Assumptions(Me.AssumptionIndex) Is Nothing Then
                Me.Struct.Assumptions.Add(value)
            Else
                Dim selAsm As Assumption = Me.Struct.Assumptions(Me.AssumptionIndex)
                selAsm = value
            End If
        End Set
    End Property

    Public ReadOnly Property AsmSort() As String
        Get
            Dim strSort As String = Nothing
            If Me.Assumption IsNot Nothing Then
                Dim intNr As Integer = Me.Struct.Assumptions.IndexOf(Me.Assumption) + 1
                strSort = CurrentLogFrame.GetStructSortNumber(Me.Struct) & "." & intNr.ToString
            End If
            Return strSort
        End Get
    End Property

    Public Property AsmRTF() As String
        Get
            Dim strRTF As String = Nothing
            If Me.Assumption IsNot Nothing Then
                strRTF = Me.Assumption.RTF
            End If
            Return strRTF
        End Get
        Set(ByVal value As String)
            If Me.Assumption Is Nothing Then NewAssumption()
            Me.Assumption.RTF = value
        End Set
    End Property

    Private Sub NewAssumption()
        If Me.Struct Is Nothing Then NewStruct()
        Me.Struct.Assumptions.Add(New Assumption())
        Me.AssumptionIndex = Me.Struct.Assumptions.Count - 1
        Me.NewlyAdded = LogFrame.ObjectTypes.Assumption
    End Sub
#End Region 'Assumption columns

End Class



Public Class GridRows
    Inherits System.Collections.CollectionBase

    Private SearchStruct As Struct
    Private boolHideEmptyCells As Boolean
    Private boolShowInd As Boolean = True
    Private boolShowVer As Boolean = True
    Private boolShowAsm As Boolean = True
    Private boolShowGoals As Boolean = True
    Private boolShowPurposes As Boolean = True
    Private boolShowOutputs As Boolean = True
    Private boolShowActivities As Boolean = True
    Private boolShowResourcesBudget As Boolean = True
    Private StructNumber As Integer

    Public Event RowCountChanged()
    Public Event StructLoaded(ByVal sender As Object, ByVal e As StructLoadedEventArgs)

#Region "Properties"
    Public Property HideEmptyCells As Boolean
        Get
            Return boolHideEmptyCells
        End Get
        Set(ByVal value As Boolean)
            boolHideEmptyCells = value
        End Set
    End Property

    Public Property ShowIndColumn() As Boolean
        Get
            Return boolShowInd
        End Get
        Set(ByVal value As Boolean)
            boolShowInd = value
        End Set
    End Property

    Public Property ShowVerColumn() As Boolean
        Get
            Return boolShowVer
        End Get
        Set(ByVal value As Boolean)
            boolShowVer = value
        End Set
    End Property

    Public Property ShowAsmColumn() As Boolean
        Get
            Return boolShowAsm
        End Get
        Set(ByVal value As Boolean)
            boolShowAsm = value
        End Set
    End Property

    Public Property ShowGoals() As Boolean
        Get
            Return boolShowGoals
        End Get
        Set(value As Boolean)
            boolShowGoals = value
        End Set
    End Property

    Public Property ShowPurposes() As Boolean
        Get
            Return boolShowPurposes
        End Get
        Set(value As Boolean)
            boolShowPurposes = value
        End Set
    End Property

    Public Property ShowOutputs() As Boolean
        Get
            Return boolShowOutputs
        End Get
        Set(value As Boolean)
            boolShowOutputs = value
            'If boolShowOutputs = True And ShowPurposes = False Then
            '    ShowPurposes = True
            'End If
        End Set
    End Property

    Public Property ShowActivities() As Boolean
        Get
            Return boolShowActivities
        End Get
        Set(value As Boolean)
            boolShowActivities = value
            'If boolShowActivities = True And ShowOutputs = False Then
            '    ShowOutputs = True
            'End If
        End Set
    End Property

    Public Property ShowResourcesBudget() As Boolean
        Get
            Return boolShowResourcesBudget
        End Get
        Set(value As Boolean)
            boolShowResourcesBudget = value
        End Set
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal gridrow As GridRow)
        List.Add(gridrow)
    End Sub

    Public Sub AddRange(ByVal gridrows() As GridRow)
        For i = 0 To gridrows.Length - 1
            If gridrows(i) IsNot Nothing Then
                List.Add(gridrows(i))
            End If

        Next
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal gridrow As GridRow)
        List.Insert(index, gridrow)
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of grid row is not valid, cannot be removed!")
        Else
            List.RemoveAt(index)
            RaiseEvent RowCountChanged()
        End If
    End Sub

    Public Sub Remove(ByVal gridrow As GridRow)
        If List.Contains(gridrow) = False Then
            System.Windows.Forms.MessageBox.Show("Grid row not in list!")
        Else
            List.Remove(gridrow)
        End If
    End Sub

    Public Overloads Sub RemoveAt(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show("Index of grid row is not valid, cannot be removed!")
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub RemoveStruct(ByVal struct As Struct)
        If struct IsNot Nothing Then
            Dim intStartIndex As Integer = Me.GetIndexOfFirstStruct(struct, struct.Section)
            Dim intLastIndex As Integer = intStartIndex
            Dim ParentPurpose As Purpose = Nothing
            If struct.GetType Is GetType(Output) Then _
                ParentPurpose = CurrentLogFrame.GetParent(CType(struct, Output))

            If GetRowCount(struct) > 0 Then intLastIndex += GetRowCount(struct) - 1
            For i = intLastIndex To intStartIndex Step -1
                Me.RemoveAt(i)
            Next

            If struct.GetType Is GetType(Purpose) Then
                RemoveRepeatedPurpose(struct)
            ElseIf struct.GetType Is GetType(Output) Then
                RemoveRepeatedOutput(struct, ParentPurpose)
            End If
            CurrentLogFrame.RemoveStruct(struct)
        End If
    End Sub

    Default Public Property Item(ByVal index As Integer) As GridRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), GridRow)
            End If
        End Get
        Set(ByVal value As GridRow)
            List.Item(index) = value
        End Set
    End Property

    Public Function IndexOf(ByVal gridrow As GridRow) As Integer
        Return List.IndexOf(gridrow)
    End Function

    Public Function Contains(ByVal gridrow As GridRow) As Boolean
        Return List.Contains(gridrow)
    End Function
#End Region

#Region "Get row counts"
    Public Function GetRowCount(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer
        Dim intRowCount As Integer, intRowCountAsm As Integer

        If ShowAsmColumn = True Then intRowCountAsm = GetRowCountAsm(objStruct)
        If objStruct.ResourceBudget = False Then
            If ShowIndColumn = True Then intRowCount = GetRowCountIndVer(objStruct)
        Else
            If ShowIndColumn = True Then intRowCount = GetRowCountRsc(objStruct)
        End If
        If intRowCount >= intRowCountAsm Then NmbrRows = intRowCount Else NmbrRows = intRowCountAsm

        Return NmbrRows
    End Function

    Public Function GetRowCountIndVer(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer
        Dim NmbrInd As Integer
        Dim NmbrVer As Integer

        If ShowIndColumn = True Then NmbrInd = objStruct.Indicators.Count
        If ShowVerColumn = True Then NmbrVer = Me.CountVerificationSources(objStruct)

        'number of rows=(number of indicators + 1) + number of verification means
        If NmbrInd > 0 Then
            NmbrRows = NmbrInd + NmbrVer
            If HideEmptyCells = False Then NmbrRows += 1
        End If
        If NmbrRows = 0 Then NmbrRows = 1
        Return NmbrRows
    End Function

    Public Function GetRowCountRsc(ByVal objActivity As Activity) As Integer
        Dim NmbrRows As Integer
        Dim NmbrRsc As Integer

        If ShowIndColumn = True Then NmbrRsc = objActivity.Resources.Count

        'number of rows=(number of indicators + 1) + number of verification means
        If NmbrRsc > 0 Then
            NmbrRows = NmbrRsc
            If HideEmptyCells = False Then NmbrRows += 1
        End If
        If NmbrRows = 0 Then NmbrRows = 1
        Return NmbrRows
    End Function

    Public Function GetRowCountAsm(ByVal objStruct As Struct) As Integer
        Dim NmbrRows As Integer
        If objStruct.Assumptions.Count > 0 Then
            NmbrRows = objStruct.Assumptions.Count
            If HideEmptyCells = False Then NmbrRows += 1
        End If

        Return NmbrRows
    End Function

    Public Function CountIndicatorRowsFromStruct(ByVal objStruct As Struct, _
                                                 ByVal IndicatorIndex As Integer, ByVal VerificationSourceIndex As Integer, ByVal AssumptionIndex As Integer) As Integer
        Dim NmbrRows As Integer
        Dim intCorrect As Integer

        If HideEmptyCells = True Then intCorrect = 1

        If Me.GetRowCountIndVer(objStruct) >= Me.GetRowCountAsm(objStruct) Then
            Dim NmbrInd As Integer
            Dim NmbrVer As Integer
            If VerificationSourceIndex = -1 Then VerificationSourceIndex = 0
            If IndicatorIndex > 0 And IndicatorIndex <= objStruct.Indicators.Count Then
                NmbrInd = IndicatorIndex
                If ShowVerColumn = True Then
                    For i = 0 To IndicatorIndex - 1

                        If HideEmptyCells = False Then
                            NmbrVer += objStruct.Indicators(i).VerificationSources.Count
                        Else
                            If objStruct.Indicators(i).VerificationSources.Count > 1 Then
                                NmbrVer += objStruct.Indicators(i).VerificationSources.Count - 1
                            End If
                        End If
                    Next
                End If
                NmbrRows = NmbrInd + NmbrVer
            End If
            If ShowVerColumn = True Then NmbrRows += VerificationSourceIndex
        Else
            If ShowAsmColumn = True Then NmbrRows = AssumptionIndex
        End If
        Return NmbrRows
    End Function

    Public Function CountIndicatorRowsFromStruct(ByVal objStruct As Struct, ByVal IndicatorIndex As Integer) As Integer
        Dim NmbrRows As Integer
        Dim NmbrInd As Integer
        Dim NmbrVer As Integer
        If IndicatorIndex > 0 And IndicatorIndex <= objStruct.Indicators.Count Then
            NmbrInd = IndicatorIndex
            If ShowVerColumn = True Then
                For i = 0 To IndicatorIndex - 1
                    If HideEmptyCells = False Then
                        NmbrVer += objStruct.Indicators(i).VerificationSources.Count
                    Else
                        If objStruct.Indicators(i).VerificationSources.Count > 1 Then
                            NmbrVer += objStruct.Indicators(i).VerificationSources.Count - 1
                        End If
                    End If
                Next
            End If
            NmbrRows = NmbrInd + NmbrVer
        End If

        Return NmbrRows
    End Function

    Public Function CountResourceRowsFromStruct(ByVal objStruct As Struct, ByVal ResourceIndex As Integer, _
                                                ByVal AssumptionIndex As Integer) As Integer
        Dim NmbrRows As Integer
        Dim selActivity As Activity = TryCast(objStruct, Activity)

        If selActivity IsNot Nothing Then
            If Me.GetRowCountRsc(objStruct) >= Me.GetRowCountAsm(objStruct) Then
                If ResourceIndex > 0 And ResourceIndex <= selActivity.Resources.Count Then
                    NmbrRows = ResourceIndex
                End If
            Else
                If ShowAsmColumn = True Then NmbrRows = AssumptionIndex
            End If
        End If

        Return NmbrRows
    End Function

    Public Function CountResourceRowsFromStruct(ByVal objStruct As Struct, ByVal ResourceIndex As Integer) As Integer
        Dim NmbrRows As Integer
        Dim selActivity As Activity = TryCast(objStruct, Activity)

        If selActivity IsNot Nothing Then
            If ResourceIndex > 0 And ResourceIndex <= selActivity.Resources.Count Then
                NmbrRows = ResourceIndex
            End If
        End If

        Return NmbrRows
    End Function

    Public ReadOnly Property CountVerificationSources(ByVal objStruct As Struct) As Integer
        Get
            Dim NmbrInd As Integer = objStruct.Indicators.Count
            Dim NmbrVer As Integer

            If NmbrInd > 0 Then
                For Each indTemp As Indicator In objStruct.Indicators
                    If HideEmptyCells = False Then
                        NmbrVer += indTemp.VerificationSources.Count
                    Else
                        If indTemp.VerificationSources.Count > 1 Then
                            NmbrVer += indTemp.VerificationSources.Count - 1
                        End If
                    End If
                Next
            End If

            Return NmbrVer
        End Get
    End Property

    Public ReadOnly Property CountBudgetItems(ByVal objStruct As Struct) As Integer
        Get
            Dim selActivity As Activity = TryCast(objStruct, Activity)
            Dim NmbrBudgetItems As Integer

            If selActivity IsNot Nothing Then
                Dim NmbrResource As Integer = selActivity.Resources.Count

                If NmbrResource > 0 Then
                    For Each objResource As Resource In selActivity.Resources
                        NmbrBudgetItems += objResource.BudgetItemReferences.Count
                    Next
                End If
            End If

            Return NmbrBudgetItems
        End Get
    End Property
#End Region

#Region "Load data"
    Public Sub LoadSections()
        Me.Clear()
        StructNumber = 0

        If ShowGoals = False And ShowPurposes = False And ShowOutputs = False And ShowActivities = False Then
            ShowPurposes = True
            frmParent.RibbonButtonShowPurposes.Checked = True
        End If

        'Section: Goals
        If ShowGoals = True Then LoadSections_SectionGoals()

        'Section: Purpose(s)
        If ShowPurposes = True Then LoadSections_SectionPurposes()

        'Section: Outputs
        If ShowOutputs = True Then LoadSections_SectionOutputs()

        'Section: Activities
        If ShowActivities = True Then LoadSections_SectionActivities()
    End Sub

    Private Sub LoadSections_SectionGoals()
        Dim rowSectionTitle As New GridRow(1, CurrentLogFrame.SectionTitles(0))
        With rowSectionTitle
            .RowType = GridRow.RowTypes.Section
            .IndicatorIndex = 0
            .VerificationSourceIndex = 0
            .AssumptionIndex = 0
        End With
        Me.Add(rowSectionTitle)
        StructNumber += 1
        RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))

        If CurrentLogFrame.Goals.Count > 0 Then
            For Each selGoal As Goal In CurrentLogFrame.Goals
                Dim intScope As Integer = GetRowCount(selGoal) - 1
                If intScope < 0 Then intScope = 0
                Dim NewRow(intScope) As GridRow

                NewRow(0) = New GridRow(1, selGoal)
                LoadSections_Children(1, selGoal, NewRow)
                Me.AddRange(NewRow)
                StructNumber += 1
                RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))
            Next
        End If
        If Me.HideEmptyCells = False Then
            Me.Add(New GridRow(1))
        End If
    End Sub

    Private Sub LoadSections_SectionPurposes()
        Dim rowSectionTitle As New GridRow(2, CurrentLogFrame.SectionTitles(1))
        With rowSectionTitle
            .RowType = GridRow.RowTypes.Section
            .IndicatorIndex = 0
            .VerificationSourceIndex = 0
            .AssumptionIndex = 0
        End With
        Me.Add(rowSectionTitle)
        StructNumber += 1
        RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))

        If CurrentLogFrame.Purposes.Count > 0 Then
            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                Dim intScope As Integer = GetRowCount(selPurpose) - 1
                If intScope < 0 Then intScope = 0
                Dim NewRow(intScope) As GridRow

                NewRow(0) = New GridRow(2, selPurpose)
                LoadSections_Children(2, selPurpose, NewRow)

                Me.AddRange(NewRow)
                StructNumber += 1
                RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))
            Next
        End If
        If Me.HideEmptyCells = False Then
            'Dim NewPurpose As New Purpose
            'CurrentLogFrame.Purposes.Add(NewPurpose)
            Me.Add(New GridRow(2))
        End If
    End Sub

    Private Sub LoadSections_SectionOutputs()
        Dim rowSectionTitle As New GridRow(3, CurrentLogFrame.SectionTitles(2))
        With rowSectionTitle
            .RowType = GridRow.RowTypes.Section
            .IndicatorIndex = 0
            .VerificationSourceIndex = 0
            .AssumptionIndex = 0
        End With
        Me.Add(rowSectionTitle)
        StructNumber += 1
        RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))

        If CurrentLogFrame.Purposes.Count > 0 Then
            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                If CurrentLogFrame.Purposes.Count > 1 Then
                    Dim rowRepeatPurpose As New GridRow(3, -1, -1, -1, selPurpose)
                    rowRepeatPurpose.RowType = GridRow.RowTypes.RepeatPurpose
                    Me.Add(rowRepeatPurpose)
                End If
                For Each selOutput As Output In selPurpose.Outputs
                    Dim intScope As Integer = GetRowCount(selOutput) - 1
                    If intScope < 0 Then intScope = 0
                    Dim NewRow(intScope) As GridRow

                    NewRow(0) = New GridRow(3, selOutput)
                    LoadSections_Children(3, selOutput, NewRow)
                    Me.AddRange(NewRow)
                    StructNumber += 1
                    RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))
                Next
                If Me.HideEmptyCells = False Then
                    Me.Add(New GridRow(3))
                End If
            Next
        Else
            If Me.HideEmptyCells = False Then
                Dim NewOutput As New Output
                Me.Add(New GridRow(3))
            End If
        End If

    End Sub

    Private Sub LoadSections_SectionActivities()
        Dim rowSectionTitle As New GridRow(4, CurrentLogFrame.SectionTitles(3))
        With rowSectionTitle
            .RowType = GridRow.RowTypes.Section
            .IndicatorIndex = 0
            .VerificationSourceIndex = 0
            .ResourceIndex = 0
            .AssumptionIndex = 0
        End With
        Me.Add(rowSectionTitle)
        StructNumber += 1
        RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))

        If CurrentLogFrame.Purposes.Count > 0 Then
            For Each selPurpose As Purpose In CurrentLogFrame.Purposes
                If CurrentLogFrame.Purposes.Count > 1 Then
                    Dim rowRepeatPurpose As New GridRow(4, -1, -1, -1, selPurpose)
                    rowRepeatPurpose.RowType = GridRow.RowTypes.RepeatPurpose
                    Me.Add(rowRepeatPurpose)

                End If
                If selPurpose.Outputs.Count = 0 Then
                    Me.Add(New GridRow(4))
                Else
                    For Each selOutput As Output In selPurpose.Outputs
                        If selPurpose.Outputs.Count > 1 Then
                            Dim rowRepeatOutput As New GridRow(4, -1, -1, -1, selOutput)
                            rowRepeatOutput.RowType = GridRow.RowTypes.RepeatOutput
                            Me.Add(rowRepeatOutput)
                        End If
                        For Each selActivity As Activity In selOutput.Activities
                            LoadSections_SectionActivities_Activity(selActivity)

                            StructNumber += 1
                            RaiseEvent StructLoaded(Me, New StructLoadedEventArgs(StructNumber))
                        Next
                        If Me.HideEmptyCells = False Then
                            Me.Add(New GridRow(4))
                        End If
                    Next
                End If
            Next
        Else
            If Me.HideEmptyCells = False Then
                Me.Add(New GridRow(4))
            End If
        End If
    End Sub

    Private Sub LoadSections_SectionActivities_Activity(ByVal selActivity As Activity)
        Dim intScope As Integer = GetRowCount(selActivity) - 1
        If intScope < 0 Then intScope = 0
        Dim NewRow(intScope) As GridRow

        NewRow(0) = New GridRow(4, selActivity)
        If ShowResourcesBudget = True Then
            LoadSections_Resources(4, selActivity, NewRow)
        Else
            LoadSections_Children(4, selActivity, NewRow)
        End If

        Me.AddRange(NewRow)

        If selActivity.IsProcess Then
            For Each selChildActivity As Activity In selActivity.Activities
                LoadSections_SectionActivities_Activity(selChildActivity)
            Next
        End If
    End Sub

    Private Sub LoadSections_Children(ByVal intSection As Integer, ByVal selStruct As Struct, _
                                      ByVal NewRow() As GridRow)
        Dim intCount As Integer, intCountAsm As Integer
        Dim intIndIndex As Integer, intVerIndex As Integer, intAsmIndex As Integer
        If Me.ShowIndColumn = True And selStruct.Indicators.Count > 0 Then
            For Each selInd As Indicator In selStruct.Indicators
                If selStruct.Indicators.IndexOf(selInd) > 0 Then
                    intCount += 1
                    NewRow(intCount) = New GridRow(intSection, selStruct)
                End If
                intIndIndex = selStruct.Indicators.IndexOf(selInd)
                NewRow(intCount).IndicatorIndex = intIndIndex

                If Me.ShowVerColumn = True And selInd.VerificationSources.Count > 0 Then
                    For Each selVer As VerificationSource In selInd.VerificationSources
                        If selInd.VerificationSources.IndexOf(selVer) > 0 Then
                            intCount += 1
                            NewRow(intCount) = New GridRow(intSection, selStruct)
                            NewRow(intCount).IndicatorIndex = intIndIndex
                        End If
                        intVerIndex = selInd.VerificationSources.IndexOf(selVer)
                        NewRow(intCount).VerificationSourceIndex = intVerIndex
                    Next
                    'add empty row for new verification source
                    If Me.HideEmptyCells = False Then
                        intCount += 1
                        intVerIndex += 1
                        NewRow(intCount) = New GridRow(intSection, selStruct)
                        NewRow(intCount).IndicatorIndex = intIndIndex
                        NewRow(intCount).VerificationSourceIndex = intVerIndex

                    End If
                End If
            Next
            'add empty row for new indicator
            If Me.HideEmptyCells = False Then
                intCount += 1
                intIndIndex += 1
                NewRow(intCount) = New GridRow(intSection, selStruct)
                NewRow(intCount).IndicatorIndex = intIndIndex
            End If
        End If

        If Me.ShowAsmColumn = True And selStruct.Assumptions.Count > 0 Then
            For Each selAsm As Assumption In selStruct.Assumptions
                If selStruct.Assumptions.IndexOf(selAsm) >= GetRowCountIndVer(selStruct) Then _
                    NewRow(intCountAsm) = New GridRow(intSection, selStruct)
                intAsmIndex = selStruct.Assumptions.IndexOf(selAsm)
                NewRow(intCountAsm).AssumptionIndex = intAsmIndex
                intCountAsm += 1
            Next

            'add empty row for new assumption
            If Me.HideEmptyCells = False Then
                Dim boolAddNewRow As Boolean
                If Me.ShowIndColumn = True Then
                    If intCountAsm > intCount Then boolAddNewRow = True
                Else
                    boolAddNewRow = True
                End If

                If boolAddNewRow = True Then
                    If Me.HideEmptyCells = False Then
                        intAsmIndex += 1
                        NewRow(intCountAsm) = New GridRow(intSection, selStruct)
                        NewRow(intCountAsm).AssumptionIndex = intAsmIndex
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub LoadSections_Resources(ByVal intSection As Integer, ByVal selStruct As Struct, ByVal NewRow() As GridRow)
        Dim intCount As Integer, intCountAsm As Integer
        Dim intRscIndex As Integer, intBudIndex As Integer, intAsmIndex As Integer
        Dim selActivity As Activity = TryCast(selStruct, Activity)

        If selActivity IsNot Nothing Then Exit Sub
        If Me.ShowIndColumn = True And selActivity.Resources.Count > 0 Then
            For Each selRsc As Resource In selActivity.Resources
                If selActivity.Resources.IndexOf(selRsc) > 0 Then NewRow(intCount) = New GridRow(intSection, selActivity)
                intRscIndex = selActivity.Resources.IndexOf(selRsc)
                NewRow(intCount).ResourceIndex = intRscIndex

                If Me.ShowVerColumn = True Then

                    'add empty row for new budget item
                    If Me.HideEmptyCells = False Then
                        intBudIndex += 1
                        NewRow(intCount) = New GridRow(intSection, selActivity)
                        NewRow(intCount).ResourceIndex = intRscIndex
                    End If
                End If
                intCount += 1
            Next
            'add empty row for new resource
            If Me.HideEmptyCells = False Then
                intRscIndex += 1
                NewRow(intCount) = New GridRow(intSection, selActivity)
                NewRow(intCount).ResourceIndex = intRscIndex
            End If
        Else
            intCount = -1
        End If

        If Me.ShowAsmColumn = True And selActivity.Assumptions.Count > 0 Then
            For Each selAsm As Assumption In selActivity.Assumptions
                If selActivity.Assumptions.IndexOf(selAsm) > intCount Then _
                    NewRow(intCountAsm) = New GridRow(intSection, selActivity)
                intAsmIndex = selActivity.Assumptions.IndexOf(selAsm)
                NewRow(intCountAsm).AssumptionIndex = intAsmIndex
                intCountAsm += 1
            Next

            'add empty row for new assumption
            Dim boolAddNewRow As Boolean
            If Me.ShowIndColumn = True Then
                If intCountAsm > intCount Then boolAddNewRow = True
            Else
                boolAddNewRow = True
            End If

            If boolAddNewRow = True Then
                If Me.HideEmptyCells = False Then
                    intAsmIndex += 1
                    NewRow(intCountAsm) = New GridRow(intSection, selActivity)
                    NewRow(intCountAsm).AssumptionIndex = intAsmIndex
                End If
            End If
        End If
    End Sub
#End Region 'Load data

#Region "Add to logframe"
    Protected Overrides Sub OnInsert(ByVal index As Integer, ByVal value As Object)
        MyBase.OnInsert(index, value)
    End Sub

    Protected Overrides Sub OnSetComplete(ByVal index As Integer, ByVal oldValue As Object, ByVal newValue As Object)
        Dim selGridRow As GridRow = CType(newValue, GridRow)
        MyBase.OnSetComplete(index, oldValue, newValue)

        AddToLogFrame(selGridRow, index)
    End Sub

    Public Sub AddToLogFrame(ByVal selGridRow As GridRow, ByVal index As Integer)
        If selGridRow.NewlyAdded <> 0 And Me.HideEmptyCells = False Then
            Dim intRowIndex As Integer = index + 1

            Select Case selGridRow.NewlyAdded

                Case LogFrame.ObjectTypes.Goal
                    AddGoal(selGridRow, intRowIndex)
                Case LogFrame.ObjectTypes.Purpose, LogFrame.ObjectTypes.PurposeHidden
                    AddPurpose(selGridRow, intRowIndex)
                Case LogFrame.ObjectTypes.Output, LogFrame.ObjectTypes.OutputHidden
                    AddOutput(selGridRow, index, intRowIndex)
                Case LogFrame.ObjectTypes.Activity
                    AddActivity(selGridRow, index, intRowIndex)
                Case LogFrame.ObjectTypes.Indicator
                    AddIndicator(selGridRow, index, intRowIndex)
                Case LogFrame.ObjectTypes.VerificationSource
                    AddVerificationSource(selGridRow, index, intRowIndex)
                Case LogFrame.ObjectTypes.Resource
                    AddResource(selGridRow, index, intRowIndex)
                Case LogFrame.ObjectTypes.Assumption
                    AddAssumption(selGridRow, index, intRowIndex)
            End Select
            RaiseEvent RowCountChanged()

            selGridRow.NewlyAdded = 0
        End If
    End Sub

    Private Sub AddGoal(ByVal selGridRow As GridRow, ByVal intRowIndex As Integer)
        Dim selGoal As Goal = CType(selGridRow.Struct, Goal)
        CurrentLogFrame.Goals.Add(selGoal)

        'CurrentUndoList.InsertNewOperation(selGoal, CurrentLogFrame.Goals, , selGridRow.StructSort, True)
    End Sub

    Private Sub AddPurpose(ByVal selGridRow As GridRow, ByVal intRowIndex As Integer)
        Dim selPurpose As Purpose = CType(selGridRow.Struct, Purpose)
        If selGridRow.NewlyAdded <> LogFrame.ObjectTypes.PurposeHidden Then
            CurrentLogFrame.Purposes.Add(selPurpose)
        End If

        'CurrentUndoList.InsertNewOperation(selPurpose, CurrentLogFrame.Purposes, , selGridRow.StructSort, True)
    End Sub

    Private Sub AddOutput(ByVal selGridRow As GridRow, ByVal index As Integer, ByVal intRowIndex As Integer)
        Dim selOutput As Output = CType(selGridRow.Struct, Output)
        Dim PreviousRow As GridRow = CType(Me(index - 1), GridRow)
        Dim ParentPurpose As Purpose = Nothing

        If PreviousRow.RowType = GridRow.RowTypes.Section Then
            ParentPurpose = CurrentLogFrame.Purposes(0)
        Else
            If PreviousRow.Struct.GetType Is GetType(Output) Then
                ParentPurpose = CurrentLogFrame.GetParent(CType(PreviousRow.Struct, Output))
            ElseIf PreviousRow.Struct.GetType Is GetType(Purpose) Then
                ParentPurpose = CType(PreviousRow.Struct, Purpose)
            End If
        End If
        If ParentPurpose Is Nothing Then ParentPurpose = HiddenPurpose()

        If selGridRow.NewlyAdded <> LogFrame.ObjectTypes.OutputHidden Then
            ParentPurpose.Outputs.Add(selOutput)
        End If

        'CurrentUndoList.InsertNewOperation(selOutput, ParentPurpose.Outputs, , selGridRow.StructSort, True)
    End Sub

    Private Sub AddActivity(ByVal selGridRow As GridRow, ByVal index As Integer, ByVal intRowIndex As Integer)
        Dim selActivity As Activity = CType(selGridRow.Struct, Activity)
        Dim PreviousRow As GridRow = CType(Me(index - 1), GridRow)
        Dim ParentPurpose As Purpose = Nothing
        Dim ParentOutput As Output = Nothing
        If PreviousRow.RowType = GridRow.RowTypes.Section Then
            ParentPurpose = CurrentLogFrame.Purposes(0)
            If ParentPurpose Is Nothing Then ParentPurpose = HiddenPurpose()
            ParentOutput = ParentPurpose.Outputs(0)
            If ParentOutput Is Nothing Then ParentOutput = HiddenOutput(ParentPurpose)
        Else
            If PreviousRow.Struct.GetType Is GetType(Activity) Then
                ParentOutput = CurrentLogFrame.GetParent(CType(PreviousRow.Struct, Activity))
            ElseIf PreviousRow.Struct.GetType Is GetType(Output) Then
                ParentOutput = CType(PreviousRow.Struct, Output)
            ElseIf PreviousRow.Struct.GetType Is GetType(Purpose) Then
                ParentPurpose = CType(PreviousRow.Struct, Purpose)
                ParentOutput = ParentPurpose.Outputs(0)
                If ParentOutput Is Nothing Then ParentOutput = HiddenOutput(ParentPurpose)
            End If
        End If

        ParentOutput.Activities.AddToProcess(selActivity)

        'CurrentUndoList.InsertNewOperation(selActivity, ParentOutput.Activities, , selGridRow.StructSort, True)
    End Sub

    Private Sub AddHidden(ByVal selGridRow As GridRow, ByVal index As Integer, ByVal intRowIndex As Integer)
        Dim selStruct As Struct = selGridRow.Struct

        Select Case selGridRow.Section
            Case LogFrame.SectionTypes.GoalsSection
                If CurrentLogFrame.Goals.Contains(selStruct) = False Then
                    AddGoal(selGridRow, intRowIndex)
                End If
            Case LogFrame.SectionTypes.PurposesSection
                If CurrentLogFrame.Purposes.Contains(selStruct) = False Then
                    AddPurpose(selGridRow, intRowIndex)
                End If
            Case LogFrame.SectionTypes.OutputsSection
                Dim ParentPurpose As Purpose = CurrentLogFrame.GetParent(CType(selStruct, Output))
                If ParentPurpose Is Nothing Then
                    AddOutput(selGridRow, index, intRowIndex)
                End If
            Case LogFrame.SectionTypes.ActivitiesSection
                Dim ParentOutput As Output = CurrentLogFrame.GetParent(CType(selStruct, Activity))
                If ParentOutput Is Nothing Then
                    AddActivity(selGridRow, index, intRowIndex)
                End If
        End Select
    End Sub

    Private Sub AddIndicator(ByVal selGridRow As GridRow, ByVal index As Integer, ByVal intRowIndex As Integer)
        Dim selStruct As Struct = selGridRow.Struct
        If selStruct.RTF = Nothing Then AddHidden(selGridRow, index, intRowIndex)

        'CurrentUndoList.InsertNewOperation(selGridRow.Indicator, selStruct.Indicators, , selGridRow.IndSort, True)
    End Sub

    Private Sub AddVerificationSource(ByVal selGridRow As GridRow, ByVal index As Integer, _
                                      ByVal intRowIndex As Integer)
        Dim selStruct As Struct = selGridRow.Struct
        If selStruct.RTF = Nothing Then AddHidden(selGridRow, index, intRowIndex)
        Dim selIndicator As Indicator = selGridRow.Indicator

        'CurrentUndoList.InsertNewOperation(selGridRow.VerificationSource, selIndicator.VerificationSources, , selGridRow.VerSort, True)
    End Sub

    Private Sub AddResource(ByVal selGridRow As GridRow, ByVal index As Integer, ByVal intRowIndex As Integer)
        Dim selActivity As Activity = TryCast(selGridRow.Struct, Activity)

        If selActivity IsNot Nothing Then
            If selActivity.RTF = Nothing Then AddHidden(selGridRow, index, intRowIndex)

            'CurrentUndoList.InsertNewOperation(selGridRow.Resource, selActivity.Resources, , selGridRow.RscSort, True)
        End If
    End Sub

    Private Sub AddAssumption(ByVal selGridRow As GridRow, ByVal index As Integer, _
                              ByVal intRowIndex As Integer)
        Dim selStruct As Struct = selGridRow.Struct
        If selStruct.RTF = Nothing Then AddHidden(selGridRow, index, intRowIndex)

        'CurrentUndoList.InsertNewOperation(selGridRow.Assumption, selStruct.Assumptions, , selGridRow.AsmSort, True)
    End Sub
#End Region

#Region "Manage repeated rows"
    Private Sub RemoveRepeatedPurpose(ByVal RemovedPurpose As Purpose)
        Dim selGridRow As GridRow
        Dim intOldSection As Integer = 4

        If CurrentLogFrame.Purposes.Count = 2 Then
            For i = Me.Count - 1 To 0 Step -1
                selGridRow = Me.Item(i)
                intOldSection = Me.Item(i).Section
                If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose Then
                    If selGridRow.Struct IsNot CurrentLogFrame.Purposes(0) Then
                        Me.RemoveAt(i + 1)
                    End If

                    Me.RemoveAt(i)
                End If

            Next
        End If
    End Sub

    Private Sub RemoveRepeatedOutput(ByVal RemovedOutput As Output, ByVal Parentpurpose As Purpose)
        Dim selGridRow As GridRow
        Dim boolLastEmptyRow As Boolean = True

        If Parentpurpose.Outputs.Count = 2 Then
            For i = Me.Count - 1 To 0 Step -1
                selGridRow = Me.Item(i)

                If selGridRow.RowType = GridRow.RowTypes.RepeatOutput AndAlso CurrentLogFrame.GetParent(CType(selGridRow.Struct, Output)) Is Parentpurpose Then
                    If boolLastEmptyRow = True Then
                        Me.RemoveAt(i + 1)
                        boolLastEmptyRow = False
                    End If
                    Me.RemoveAt(i)
                End If
            Next
        End If
    End Sub

    Private Function HiddenPurpose() As Purpose
        Dim ParentPurpose As Purpose

        CurrentLogFrame.Purposes.Add(New Purpose())
        ParentPurpose = CurrentLogFrame.Purposes(0)

        Dim lstSectionPurposes As List(Of GridRow) = Me.GetSection(LogFrame.SectionTypes.PurposesSection)
        lstSectionPurposes(0).Struct = CType(ParentPurpose, Purpose)

        Return ParentPurpose
    End Function

    Private Function HiddenOutput(ByVal ParentPurpose As Purpose) As Output
        Dim ParentOutput As Output

        ParentPurpose.Outputs.Add(New Output())
        ParentOutput = ParentPurpose.Outputs(0)

        Dim lstSectionOutputs As List(Of GridRow) = Me.GetSection(LogFrame.SectionTypes.OutputsSection)
        lstSectionOutputs(0).Struct = CType(ParentOutput, Output)

        Return ParentOutput
    End Function
#End Region

#Region "Get"
    Public Function GetIndexOfFirstStruct(ByVal SearchStruct As Struct) As Integer
        Dim selRow As GridRow = Nothing
        Dim index As Integer
        For index = 0 To List.Count - 1
            selRow = CType(List(index), GridRow)
            If selRow.Struct Is SearchStruct Then Exit For
        Next
        Return index
    End Function

    Public Function GetIndexOfFirstStruct(ByVal SearchStruct As Struct, ByVal intSearchSection As Integer) As Integer
        Dim index As Integer = -1
        If SearchStruct IsNot Nothing Then
            Dim lstSection As List(Of GridRow) = GetSection(intSearchSection)
            Dim selRow As GridRow = Nothing

            For i = 0 To lstSection.Count - 1
                selRow = CType(lstSection(i), GridRow)
                If selRow.Struct Is SearchStruct Then
                    index = List.IndexOf(selRow)
                    Exit For
                End If
            Next
        End If
        Return index
    End Function

    Private Function GetAllStruct() As List(Of GridRow)
        Dim lstFound As New List(Of GridRow)
        For Each selGridRow As GridRow In List
            If selGridRow.Struct Is SearchStruct Then lstFound.Add(selGridRow)
        Next
        Return lstFound
    End Function

    Public Function GetSection(ByVal SearchSection As Integer) As List(Of GridRow)
        Dim lstFound As New List(Of GridRow)
        For i = 0 To List.Count - 1
            Dim selGridRow As GridRow = CType(Me(i), GridRow)
            If selGridRow.Section = SearchSection And selGridRow.RowType <> GridRow.RowTypes.Section Then lstFound.Add(selGridRow)
        Next
        Return lstFound
    End Function

    Private Function GetRepeatedPurposes() As List(Of GridRow)
        Dim lstSectionActivities As List(Of GridRow) = GetSection(LogFrame.SectionTypes.ActivitiesSection)
        Dim lstFound As New List(Of GridRow)
        For i = 0 To lstSectionActivities.Count - 1
            Dim selGridRow As GridRow = CType(lstSectionActivities(i), GridRow)
            If selGridRow.RowType = GridRow.RowTypes.RepeatPurpose Then lstFound.Add(selGridRow)
        Next
        Return lstFound
    End Function
#End Region
End Class

Public Class StructLoadedEventArgs
    Inherits EventArgs

    Public Property StructNumber As Integer

    Public Sub New(ByVal intStructNumber As Integer)
        MyBase.New()

        Me.StructNumber = intStructNumber
    End Sub
End Class
