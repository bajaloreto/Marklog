Partial Public Class DataGridViewLogframe
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As LogframeRow
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name

        If e.RowIndex > Me.Grid.Count - 1 Then Exit Sub

        ' Store a reference to the LogframeRow for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), LogframeRow)
        End If

        If RowTmp IsNot Nothing Then
            ' Set the cell value to paint using the LogframeRow retrieved.
            Select Case strColName
                Case "StructSort"
                    e.Value = RowTmp.StructSort
                Case "StructRTF"
                    e.Value = RowTmp.StructRtf
                Case "IndSort"
                    If IsResourceBudgetRow(RowTmp) = False Then
                        e.Value = RowTmp.IndicatorSort
                    Else
                        e.Value = RowTmp.ResourceSort
                    End If
                Case "IndRTF"
                    If IsResourceBudgetRow(RowTmp) = False Then
                        e.Value = RowTmp.IndicatorRtf
                    Else
                        e.Value = RowTmp.ResourceRtf
                    End If
                Case "VerSort"
                    If IsResourceBudgetRow(RowTmp) = False Then
                        e.Value = RowTmp.VerificationSourceSort
                    End If
                Case "VerRTF"
                    If IsResourceBudgetRow(RowTmp) = False Then
                        e.Value = RowTmp.VerificationSourceRtf
                    Else
                        e.Value = RowTmp.TotalCostAmount
                    End If
                Case "AsmSort"
                    e.Value = RowTmp.AssumptionSort
                Case "AsmRTF"
                    e.Value = RowTmp.AssumptionRtf
            End Select
        End If
        MyBase.OnCellValueNeeded(e)
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As LogframeRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim intCellValue As Integer
        Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)
        Dim selRowIndex As Integer = e.RowIndex
        Dim intSection As Integer = Me.Grid(selRowIndex - 1).Section
        Dim sngAmount As Single

        'necessary for pasting into non-active cell
        If ctlRTF Is Nothing Then
            BeginEdit(False)
            ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
        End If

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new LogframeRow object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As LogframeRow = CType(Me.Grid(selRowIndex), LogframeRow)

                If IsResourceBudgetRow(CurrentGridRow) = False Then
                    Me.EditRow = New LogframeRow(CurrentGridRow.Section, CurrentGridRow.Struct, CurrentGridRow.Indicator, CurrentGridRow.VerificationSource, CurrentGridRow.Assumption)
                Else
                    Me.EditRow = New LogframeRow(CType(CurrentGridRow.Struct, Activity), CurrentGridRow.Resource, CurrentGridRow.Assumption)
                End If
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        ' Set the appropriate objRowEdit property to the cell value entered.
        strColName = Me.Columns(e.ColumnIndex).Name
        Select Case strColName
            Case "StructRTF", "IndRTF", "VerRTF", "AsmRTF"

                If strColName = "VerRTF" And IsResourceBudgetRow(RowTmp) Then
                    Dim strTmp As String = TryCast(ctlRTF.Text, String)

                    Single.TryParse(strTmp, sngAmount)
                Else
                    strCellValue = TryCast(ctlRTF.Rtf, String)
                End If

            Case "StructMerge", "IndMerge", "StructHeight", "IndHeight", "VerHeight", "AsmHeight"
                intCellValue = e.Value
        End Select

        Select Case strColName
            Case "StructRTF"
                If RowTmp.Struct Is Nothing Then RowTmp.Struct = InitialiseStruct(RowTmp, selRowIndex)
                RowTmp.StructRtf = strCellValue
                RowTmp.StructImageDirty = True
            Case "IndRTF"
                If IsResourceBudgetRow(RowTmp) = False Then
                    If RowTmp.Indicator Is Nothing Then RowTmp.Indicator = InitialiseIndicator(RowTmp, selRowIndex)
                    RowTmp.IndicatorRtf = strCellValue
                    RowTmp.IndicatorImageDirty = True
                Else
                    If RowTmp.Resource Is Nothing Then RowTmp.Resource = InitialiseResource(RowTmp, selRowIndex)
                    RowTmp.ResourceRtf = strCellValue
                    RowTmp.ResourceImageDirty = True
                End If
            Case "VerRTF"
                If IsResourceBudgetRow(RowTmp) = False Then
                    If RowTmp.VerificationSource Is Nothing Then RowTmp.VerificationSource = InitialiseVerificationSource(RowTmp, selRowIndex)
                    RowTmp.VerificationSourceRtf = strCellValue
                    RowTmp.VerificationSourceImageDirty = True
                Else
                    If RowTmp.Resource Is Nothing Then RowTmp.Resource = InitialiseResource(RowTmp, selRowIndex)
                    RowTmp.TotalCostAmount = sngAmount

                End If
            Case "AsmRTF"
                If RowTmp.Assumption Is Nothing Then RowTmp.Assumption = InitialiseAssumption(RowTmp, selRowIndex)
                RowTmp.AssumptionRtf = strCellValue
                RowTmp.AssumptionImageDirty = True
        End Select
    End Sub

    Private Function InitialiseStruct(ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer) As Struct
        Dim NewStruct As Struct = Nothing

        Select Case selGridRow.Section
            Case Logframe.SectionTypes.GoalsSection
                NewStruct = New Goal
                Me.Logframe.Goals.Add(NewStruct)

                UndoRedo.ItemInserted(NewStruct, Me.Logframe.Goals)
            Case Logframe.SectionTypes.PurposesSection
                NewStruct = New Purpose
                Me.Logframe.Purposes.Add(NewStruct)

                UndoRedo.ItemInserted(NewStruct, Me.Logframe.Purposes)
            Case Logframe.SectionTypes.OutputsSection
                Dim PreviousOutput As Output = TryCast(Grid.GetPreviousStruct(intRowIndex), Output)
                Dim ParentPurpose As Purpose = Nothing
                NewStruct = New Output

                If PreviousOutput.ParentPurposeGuid = Guid.Empty Then
                    If Logframe.Purposes.Count = 0 Then
                        ParentPurpose = New Purpose
                        Logframe.Purposes.Add(ParentPurpose)

                        UndoRedo.ItemInserted(ParentPurpose, Me.Logframe.Purposes)
                    Else
                        ParentPurpose = Logframe.Purposes(0)
                    End If
                Else
                    ParentPurpose = Logframe.GetPurposeByGuid(PreviousOutput.ParentPurposeGuid)
                End If

                If ParentPurpose IsNot Nothing Then
                    ParentPurpose.Outputs.Add(NewStruct)
                    UndoRedo.ItemInserted(NewStruct, ParentPurpose.Outputs)
                End If
            Case Logframe.SectionTypes.ActivitiesSection
                Dim PreviousStruct As Struct = Grid.GetPreviousStruct(intRowIndex)
                Dim ParentPurpose As Purpose = Nothing
                Dim ParentOutput As Output = Nothing
                NewStruct = New Activity

                Select Case PreviousStruct.GetType
                    Case GetType(Purpose)
                        ParentPurpose = DirectCast(PreviousStruct, Purpose)

                        If ParentPurpose.Outputs.Count = 0 Then
                            ParentPurpose.Outputs.Add(New Output)
                            UndoRedo.ItemInserted(ParentPurpose.Outputs(0), ParentPurpose.Outputs)
                        End If

                        ParentOutput = ParentPurpose.Outputs(0)
                        ParentOutput.Activities.AddToProcess(NewStruct)
                    Case GetType(Output)
                        ParentOutput = DirectCast(PreviousStruct, Output)

                        ParentOutput.Activities.AddToProcess(NewStruct)
                        UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
                    Case GetType(Activity)
                        Dim PreviousActivity As Activity = DirectCast(PreviousStruct, Activity)

                        If PreviousActivity.ParentOutputGuid = Guid.Empty And PreviousActivity.ParentActivityGuid = Guid.Empty Then
                            If Me.Logframe.Purposes.Count = 0 Then
                                ParentPurpose = New Purpose
                                Logframe.Purposes.Add(ParentPurpose)
                                UndoRedo.ItemInserted(ParentPurpose, Me.Logframe.Purposes)
                            Else
                                ParentPurpose = Logframe.Purposes(0)
                            End If
                            If ParentPurpose.Outputs.Count = 0 Then
                                ParentOutput = New Output
                                ParentPurpose.Outputs.Add(ParentOutput)
                                UndoRedo.ItemInserted(ParentOutput, ParentPurpose.Outputs)
                            Else
                                ParentOutput = ParentPurpose.Outputs(0)
                            End If

                            ParentOutput.Activities.AddToProcess(NewStruct)
                            UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
                        Else
                            If PreviousActivity.ParentOutputGuid <> Guid.Empty Then
                                ParentOutput = Logframe.GetOutputByGuid(PreviousActivity.ParentOutputGuid)
                                If ParentOutput IsNot Nothing Then
                                    ParentOutput.Activities.AddToProcess(NewStruct)
                                    UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
                                End If
                            ElseIf PreviousActivity.ParentActivityGuid <> Guid.Empty Then
                                Dim ParentActivity As Activity
                                ParentActivity = Logframe.GetActivityByGuid(PreviousActivity.ParentActivityGuid)
                                If ParentActivity IsNot Nothing Then
                                    ParentActivity.Activities.AddToProcess(NewStruct)
                                    UndoRedo.ItemInserted(NewStruct, ParentActivity.Activities)
                                End If
                            End If
                        End If
                End Select
        End Select

        Return NewStruct
    End Function

    Private Function InitialiseIndicator(ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer) As Indicator
        Dim NewIndicator As New Indicator
        Dim ParentStruct As Struct = selGridRow.Struct
        Dim PreviousIndicator As Indicator = Grid.GetPreviousIndicator(intRowIndex, selGridRow.Section)
        Dim ParentIndicators As Indicators

        If ParentStruct IsNot Nothing Then
            ParentStruct.Indicators.Add(NewIndicator)
            UndoRedo.ItemInserted(NewIndicator, ParentStruct.Indicators)
        ElseIf PreviousIndicator IsNot Nothing Then
            ParentIndicators = Me.Logframe.GetParentCollection_LogframeObject(PreviousIndicator)
            ParentIndicators.Add(NewIndicator)
            UndoRedo.ItemInserted(NewIndicator, ParentIndicators)
        Else
            ParentStruct = Grid.GetPreviousStruct(intRowIndex, selGridRow.Section)

            If ParentStruct Is Nothing Then ParentStruct = InitialiseStruct(selGridRow, intRowIndex)
            ParentStruct.Indicators.Add(NewIndicator)
            UndoRedo.ItemInserted(NewIndicator, ParentStruct.Indicators)
        End If

        Return NewIndicator
    End Function

    Private Function InitialiseResource(ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer) As Resource
        Dim NewResource As New Resource
        Dim PreviousRow As LogframeRow = Grid(intRowIndex - 1)
        Dim ParentActivity As Activity

        If PreviousRow.Resource Is Nothing Or PreviousRow.RowType <> LogframeRow.RowTypes.Normal Then
            ParentActivity = selGridRow.Struct
        Else
            ParentActivity = Logframe.GetActivityByGuid(PreviousRow.Resource.ParentStructGuid)
        End If
        If ParentActivity Is Nothing Then
            ParentActivity = InitialiseStruct(selGridRow, intRowIndex)
        End If
        ParentActivity.Resources.Add(NewResource)
        UndoRedo.ItemInserted(NewResource, ParentActivity.Resources)

        Return NewResource
    End Function

    Private Function InitialiseVerificationSource(ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer) As VerificationSource
        Dim NewVerificationSource As New VerificationSource
        Dim PreviousRow As LogframeRow = Grid(intRowIndex - 1)
        Dim ParentIndicator As Indicator

        If PreviousRow.VerificationSource Is Nothing Or PreviousRow.RowType <> LogframeRow.RowTypes.Normal Then
            ParentIndicator = selGridRow.Indicator
        Else
            ParentIndicator = Logframe.GetIndicatorByGuid(PreviousRow.VerificationSource.ParentIndicatorGuid)
        End If
        If ParentIndicator Is Nothing Then
            ParentIndicator = InitialiseIndicator(selGridRow, intRowIndex)
        End If
        ParentIndicator.VerificationSources.Add(NewVerificationSource)
        UndoRedo.ItemInserted(NewVerificationSource, ParentIndicator.VerificationSources)

        Return NewVerificationSource
    End Function

    Private Function InitialiseAssumption(ByVal selGridRow As LogframeRow, ByVal intRowIndex As Integer) As Assumption
        Dim NewAssumption As New Assumption
        Dim PreviousRow As LogframeRow = Grid(intRowIndex - 1)
        Dim ParentStruct As Struct

        If PreviousRow.Assumption Is Nothing Or PreviousRow.RowType <> LogframeRow.RowTypes.Normal Then
            ParentStruct = selGridRow.Struct
        Else
            ParentStruct = Logframe.GetStructByGuid(PreviousRow.Assumption.ParentStructGuid)
        End If
        If ParentStruct Is Nothing Then
            ParentStruct = InitialiseStruct(selGridRow, intRowIndex)
        End If
        ParentStruct.Assumptions.Add(NewAssumption)
        UndoRedo.ItemInserted(NewAssumption, ParentStruct.Assumptions)

        Return NewAssumption
    End Function

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New LogframeRow(Logframe.SectionTypes.ActivitiesSection)
        Else
            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding logframe row.
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        End If
        Me.Reload()
        MyBase.OnCancelRowEdit(e)

    End Sub

    Protected Overrides Sub OnRowDirtyStateNeeded(ByVal e As System.Windows.Forms.QuestionEventArgs)
        MyBase.OnRowDirtyStateNeeded(e)
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.IsCurrentCellDirty

        End If
    End Sub

    Protected Overrides Sub OnRowValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' Save row changes if any were made and release the edited 
        ' logframe row if there is one.
        If e.RowIndex >= Me.Grid.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then

            ' Add the new logframer object to grid.
            Me.Grid.Add(Me.EditRow)

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified logframe row in grid.
            Me.Grid(e.RowIndex) = Me.EditRow
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf Me.ContainsFocus Then

            Me.EditRow = Nothing
            Me.EditRowFlag = -1

        End If
        MyBase.OnRowValidated(e)
    End Sub
End Class
