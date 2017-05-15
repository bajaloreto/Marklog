Partial Public Class DataGridViewPlanning
    Protected Overrides Sub OnCellValueNeeded(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        MyBase.OnCellValueNeeded(e)
        Dim RowTmp As PlanningGridRow = Nothing
        Dim strColName As String = Me.Columns(e.ColumnIndex).Name
        Dim selActivity As Activity

        If e.RowIndex = RowCount - 1 Then
            Return
        End If

        ' Store a reference to the planning grid row for the row being painted.
        If e.RowIndex = EditRowFlag Then
            RowTmp = Me.EditRow
        Else
            RowTmp = CType(Me.Grid(e.RowIndex), PlanningGridRow)
        End If
        If RowTmp Is Nothing Then Exit Sub

        ' Set the cell value to paint using the Customer object retrieved.
        Select Case strColName
            Case "SortNumber"
                e.Value = RowTmp.SortNumber
            Case "RTF"
                If RowTmp.IsKeyMoment = True Then
                    If RowTmp.KeyMoment IsNot Nothing Then e.Value = RowTmp.KeyMoment.Description
                Else
                    If RowTmp.Struct IsNot Nothing Then e.Value = RowTmp.Struct.RTF
                End If
            Case "StartDate"
                If RowTmp.IsKeyMoment = True Then
                    e.Value = RowTmp.KeyMoment.ExactDateKeyMoment
                Else
                    selActivity = TryCast(RowTmp.Struct, Activity)
                    If selActivity IsNot Nothing AndAlso selActivity.ExactStartDate > Date.MinValue Then e.Value = selActivity.ExactStartDate
                End If
            Case "EndDate"
                If RowTmp.IsKeyMoment = False Then
                    selActivity = TryCast(RowTmp.Struct, Activity)
                    If selActivity IsNot Nothing AndAlso selActivity.ExactEndDate > Date.MinValue Then e.Value = selActivity.ExactEndDate
                End If
        End Select
    End Sub

    Protected Overrides Sub OnCellValuePushed(ByVal e As System.Windows.Forms.DataGridViewCellValueEventArgs)
        Dim RowTmp As PlanningGridRow
        Dim strColName As String
        Dim strCellValue As String = String.Empty
        Dim selRowIndex As Integer = e.RowIndex

        If selRowIndex < Me.Grid.Count Then
            'If the user is editing a new row, create a new planning grid row object.
            If Me.EditRow Is Nothing Then
                Dim CurrentGridRow As PlanningGridRow = CType(Me.Grid(selRowIndex), PlanningGridRow)

                Me.EditRow = New PlanningGridRow
                With EditRow
                    .KeyMoment = CurrentGridRow.KeyMoment
                    .RowType = CurrentGridRow.RowType
                    .Struct = CurrentGridRow.Struct
                    .StructHeight = CurrentGridRow.StructHeight
                    .StructImage = CurrentGridRow.StructImage
                    .StructImageDirty = CurrentGridRow.StructImageDirty
                    .Indent = CurrentGridRow.Indent
                    .SortNumber = CurrentGridRow.SortNumber
                End With
            End If
            RowTmp = Me.EditRow
            Me.EditRowFlag = e.RowIndex
        Else
            RowTmp = Me.EditRow
        End If

        ' Set the appropriate objRowEdit property to the cell value entered.
        strColName = Me.Columns(e.ColumnIndex).Name
        Select Case strColName
            Case "RTF"
                Dim ctlRTF As RichTextEditingControlLogframe = CType(Me.EditingControl, RichTextEditingControlLogframe)

                'necessary for pasting into non-active cell
                If ctlRTF Is Nothing Then
                    BeginEdit(False)
                    ctlRTF = CType(Me.EditingControl, RichTextEditingControlLogframe)
                End If

                If RowTmp.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    strCellValue = TryCast(ctlRTF.Text, String)
                ElseIf RowTmp.RowType = PlanningGridRow.RowTypes.Activity Then
                    strCellValue = TryCast(ctlRTF.Rtf, String)
                End If
        End Select

        If RowTmp.RowType = PlanningGridRow.RowTypes.KeyMoment Then
            If RowTmp.KeyMoment Is Nothing Then RowTmp.KeyMoment = InitialiseKeyMoment(RowTmp, selRowIndex)
        ElseIf RowTmp.RowType = PlanningGridRow.RowTypes.Activity Then
            If RowTmp.Struct Is Nothing Then RowTmp.Struct = InitialiseStruct(RowTmp, selRowIndex)
        End If

        Select Case strColName
            Case "RTF"
                If RowTmp.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    RowTmp.KeyMoment.Description = strCellValue
                ElseIf RowTmp.RowType = PlanningGridRow.RowTypes.Activity Then
                    RowTmp.Struct.RTF = strCellValue
                    RowTmp.StructImageDirty = True
                End If
            Case "StartDate"
                If RowTmp.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    Dim selKeyMoment As KeyMoment = RowTmp.KeyMoment
                    Dim selStartDate As Date
                    Date.TryParse(e.Value, selStartDate)

                    If selKeyMoment IsNot Nothing And selStartDate > Date.MinValue Then
                        SetStartDateByDate(selKeyMoment, selStartDate)
                    End If
                ElseIf RowTmp.RowType = PlanningGridRow.RowTypes.Activity Then
                    Dim selActivity As Activity = TryCast(RowTmp.Struct, Activity)
                    Dim selStartDate As Date
                    Date.TryParse(e.Value, selStartDate)

                    If selActivity IsNot Nothing And selStartDate > Date.MinValue Then
                        SetStartDateByDate(selActivity, selStartDate)
                    End If
                End If
            Case "EndDate"
                If RowTmp.RowType = PlanningGridRow.RowTypes.KeyMoment Then
                    Dim selKeyMoment As KeyMoment = RowTmp.KeyMoment
                    Dim selStartDate As Date
                    Date.TryParse(e.Value, selStartDate)

                    If selKeyMoment IsNot Nothing And selStartDate > Date.MinValue Then
                        SetStartDateByDate(selKeyMoment, selStartDate)
                    End If
                ElseIf RowTmp.RowType = PlanningGridRow.RowTypes.Activity Then
                    Dim selActivity As Activity = TryCast(RowTmp.Struct, Activity)
                    Dim selEndDate As Date
                    Date.TryParse(e.Value, selEndDate)

                    If selActivity IsNot Nothing And selEndDate > Date.MinValue Then
                        SetDurationByEndDate(selActivity, selEndDate)
                    End If
                End If
        End Select
    End Sub

    Private Function InitialiseStruct(ByVal selGridRow As PlanningGridRow, ByVal intRowIndex As Integer) As Struct
        Dim NewStruct As Struct = Nothing

        Dim PreviousStruct As Struct = Grid.GetPreviousStruct(intRowIndex)
        Dim ParentPurpose As Purpose = Nothing
        Dim ParentOutput As Output = Nothing
        NewStruct = New Activity

        If PreviousStruct Is Nothing Then
            'No purposes and outputs identified yet
            If CurrentLogFrame.Purposes.Count = 0 Then
                CurrentLogFrame.Purposes.Add(New Purpose())
            End If
            If CurrentLogFrame.Purposes(0).Outputs.Count = 0 Then
                CurrentLogFrame.Purposes(0).Outputs.Add(New Output())
            End If
            PreviousStruct = CurrentLogFrame.Purposes(0).Outputs(0)
        End If

        Select Case PreviousStruct.GetType
            Case GetType(Purpose)
                ParentPurpose = DirectCast(PreviousStruct, Purpose)
                ParentOutput = ParentPurpose.Outputs(0)

                ParentOutput.Activities.Add(NewStruct)
                UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
            Case GetType(Output)
                ParentOutput = DirectCast(PreviousStruct, Output)

                ParentOutput.Activities.Add(NewStruct)
                UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
            Case GetType(Activity)
                Dim PreviousActivity As Activity = DirectCast(PreviousStruct, Activity)
                'Dim ParentActivities As Activities = CurrentLogFrame.GetParentCollection(PreviousActivity)

                'If ParentActivities IsNot Nothing Then ParentActivities.Add(NewStruct)
                If PreviousActivity.ParentOutputGuid = Guid.Empty And PreviousActivity.ParentActivityGuid = Guid.Empty Then
                    If CurrentLogFrame.Purposes.Count = 0 Then
                        ParentPurpose = New Purpose
                        CurrentLogFrame.Purposes.Add(ParentPurpose)
                    Else
                        ParentPurpose = CurrentLogFrame.Purposes(0)
                    End If
                    If ParentPurpose.Outputs.Count = 0 Then
                        ParentOutput = New Output
                        ParentPurpose.Outputs.Add(ParentOutput)
                    Else
                        ParentOutput = ParentPurpose.Outputs(0)
                    End If

                    ParentOutput.Activities.Add(NewStruct)
                    UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
                Else
                    If PreviousActivity.ParentOutputGuid <> Guid.Empty Then
                        ParentOutput = CurrentLogFrame.GetOutputByGuid(PreviousActivity.ParentOutputGuid)
                        If ParentOutput IsNot Nothing Then
                            ParentOutput.Activities.Add(NewStruct)
                            UndoRedo.ItemInserted(NewStruct, ParentOutput.Activities)
                        End If

                    ElseIf PreviousActivity.ParentActivityGuid <> Guid.Empty Then
                        Dim ParentActivity As Activity
                        ParentActivity = CurrentLogFrame.GetActivityByGuid(PreviousActivity.ParentActivityGuid)
                        If ParentActivity IsNot Nothing Then
                            ParentActivity.Activities.AddToProcess(NewStruct)
                            UndoRedo.ItemInserted(NewStruct, ParentActivity.Activities)
                        End If
                    End If
                End If
        End Select

        Return NewStruct
    End Function

    Private Function InitialiseKeyMoment(ByVal selGridRow As PlanningGridRow, ByVal intRowIndex As Integer) As KeyMoment
        Dim NewKeyMoment As New KeyMoment
        Dim PreviousRow As PlanningGridRow = Grid(intRowIndex - 1)
        Dim ParentStruct As Output = Nothing
        Dim ProjectStartKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment

        With NewKeyMoment
            If ProjectStartKeyMoment IsNot Nothing Then
                .Relative = True
                .GuidReferenceMoment = ProjectStartKeyMoment.Guid
                .PeriodDirection = KeyMoment.DirectionValues.After
                .PeriodUnit = ActivityDetail.DurationUnits.Day
            Else
                .Relative = False
                .KeyMoment = Me.StartDate
            End If
        End With

        If PreviousRow Is Nothing Then
            'No purposes and outputs identified yet
            If CurrentLogFrame.Purposes.Count = 0 Then
                CurrentLogFrame.Purposes.Add(New Purpose())
            End If
            If CurrentLogFrame.Purposes(0).Outputs.Count = 0 Then
                CurrentLogFrame.Purposes(0).Outputs.Add(New Output())
                ParentStruct = CurrentLogFrame.Purposes(0).Outputs(0)
            End If
        Else
            If PreviousRow.KeyMoment Is Nothing Or PreviousRow.RowType <> PlanningGridRow.RowTypes.KeyMoment Then
                ParentStruct = TryCast(Grid.GetPreviousStruct(intRowIndex), Output)
            Else
                ParentStruct = CurrentLogFrame.GetOutputByGuid(PreviousRow.KeyMoment.ParentOutputGuid)
            End If
        End If
        If ParentStruct Is Nothing Then
            ParentStruct = InitialiseStruct(selGridRow, intRowIndex)
        End If
        ParentStruct.KeyMoments.Add(NewKeyMoment)
        UndoRedo.ItemInserted(NewKeyMoment, ParentStruct.KeyMoments)

        Return NewKeyMoment
    End Function

    Private Sub SetStartDateByDate(ByVal selKeyMoment As KeyMoment, ByVal selStartDate As Date)
        If selStartDate = Date.MinValue Then Exit Sub

        With selKeyMoment
            If .Relative = False Then
                UndoRedo.UndoBuffer_Initialise(selKeyMoment, "KeyMoment", selKeyMoment.KeyMoment)
                .KeyMoment = selStartDate
                UndoRedo.DateChanged(selStartDate)
            Else
                If .GuidReferenceMoment <> Guid.Empty Then
                    Dim selMoment As Object = CurrentLogFrame.GetReferenceMomentByGuid(.GuidReferenceMoment)
                    If selMoment Is Nothing Then Exit Sub

                    Select Case selMoment.GetType
                        Case GetType(Activity)
                            Dim RefActivity As Activity = DirectCast(selMoment, Activity)
                            Dim intNrDays As Integer

                            If .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                intNrDays = selStartDate.Subtract(RefActivity.ActivityDetail.EndDateMainActivity).Days
                            Else
                                intNrDays = RefActivity.ActivityDetail.EndDateMainActivity.Subtract(selStartDate).Days
                            End If

                            SetStartDateByDate_SetPeriod(selKeyMoment, intNrDays)
                        Case GetType(KeyMoment)
                            Dim RefKeyMoment As KeyMoment = DirectCast(selMoment, KeyMoment)
                            Dim intNrDays As Integer

                            If .PeriodDirection = KeyMoment.DirectionValues.After Then
                                intNrDays = selStartDate.Subtract(RefKeyMoment.ExactDateKeyMoment).Days
                            Else
                                intNrDays = RefKeyMoment.ExactDateKeyMoment.Subtract(selStartDate).Days
                            End If
                            
                            SetStartDateByDate_SetPeriod(selKeyMoment, intNrDays)
                    End Select
                End If
            End If
        End With
    End Sub

    Private Sub SetStartDateByDate_SetPeriod(ByVal selKeyMoment As KeyMoment, ByVal intNrDays As Integer)
        With selKeyMoment
            If intNrDays Mod 7 = 0 Then
                UndoRedo.UndoBuffer_Initialise(selKeyMoment, "Period", selKeyMoment.Period)
                .Period = intNrDays / 7
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selKeyMoment, "PeriodUnit", selKeyMoment.PeriodUnit)
                .PeriodUnit = DurationUnits.Week
                UndoRedo.OptionChanged(.PeriodUnit)
            Else
                UndoRedo.UndoBuffer_Initialise(selKeyMoment, "Period", selKeyMoment.Period)
                .Period = intNrDays
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selKeyMoment, "PeriodUnit", selKeyMoment.PeriodUnit)
                .PeriodUnit = DurationUnits.Day
                UndoRedo.OptionChanged(.PeriodUnit)
            End If
        End With
    End Sub

    Private Sub SetStartDateByDate(ByVal selActivity As Activity, ByVal selStartDate As Date)
        If selStartDate = Date.MinValue Then Exit Sub

        With selActivity.ActivityDetail
            If .Relative = False Then
                UndoRedo.UndoBuffer_Initialise(selActivity.ActivityDetail, "StartDate", selActivity.ActivityDetail.StartDate)
                .StartDate = selStartDate
                UndoRedo.DateChanged(selStartDate)
            Else
                If .GuidReferenceMoment <> Guid.Empty Then
                    Dim selMoment As Object = CurrentLogFrame.GetReferenceMomentByGuid(.GuidReferenceMoment)
                    If selMoment Is Nothing Then Exit Sub

                    Select Case selMoment.GetType
                        Case GetType(Activity)
                            Dim RefActivity As Activity = DirectCast(selMoment, Activity)
                            Dim intNrDays As Integer

                            If .PeriodDirection = ActivityDetail.DirectionValues.After Then
                                intNrDays = selStartDate.Subtract(RefActivity.ActivityDetail.EndDateMainActivity).Days
                            Else
                                intNrDays = RefActivity.ActivityDetail.EndDateMainActivity.Subtract(selStartDate).Days
                            End If
                            intNrDays += 1

                            SetStartDateByDate_SetPeriod(selActivity.ActivityDetail, intNrDays)
                        Case GetType(KeyMoment)
                            Dim RefKeyMoment As KeyMoment = DirectCast(selMoment, KeyMoment)
                            Dim intNrDays As Integer

                            If .PeriodDirection = KeyMoment.DirectionValues.After Then
                                intNrDays = selStartDate.Subtract(RefKeyMoment.ExactDateKeyMoment).Days
                            Else
                                intNrDays = RefKeyMoment.ExactDateKeyMoment.Subtract(selStartDate).Days
                            End If
                            intNrDays += 1

                            SetStartDateByDate_SetPeriod(selActivity.ActivityDetail, intNrDays)
                    End Select
                End If
            End If
        End With
    End Sub

    Private Sub SetStartDateByDate_SetPeriod(ByVal selActivityDetail As ActivityDetail, ByVal intNrDays As Integer)
        With selActivityDetail
            If intNrDays Mod 7 = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Period", selActivityDetail.Period)
                .Period = intNrDays / 7
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "PeriodUnit", selActivityDetail.PeriodUnit)
                .PeriodUnit = ActivityDetail.DurationUnits.Week
                UndoRedo.OptionChanged(.PeriodUnit)
            Else
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Period", selActivityDetail.Period)
                .Period = intNrDays
                UndoRedo.ValueChanged(.Period)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "PeriodUnit", selActivityDetail.PeriodUnit)
                .PeriodUnit = ActivityDetail.DurationUnits.Day
                UndoRedo.OptionChanged(.PeriodUnit)
            End If
        End With
    End Sub

    Private Sub SetDurationByEndDate(ByVal selActivity As Activity, ByVal selEndDate As Date)
        Dim intNrDays As Integer = selEndDate.Subtract(selActivity.ActivityDetail.StartDateMainActivity).Days + 1

        SetDurationByEndDate_SetDuration(selActivity.ActivityDetail, intNrDays)
    End Sub

    Private Sub SetDurationByEndDate_SetDuration(ByVal selActivityDetail As ActivityDetail, ByVal intNrDays As Integer)
        With selActivityDetail
            If intNrDays Mod 7 = 0 Then
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Duration", selActivityDetail.Duration)
                .Duration = intNrDays / 7
                UndoRedo.ValueChanged(.Duration)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "DurationUnit", selActivityDetail.DurationUnit)
                .DurationUnit = ActivityDetail.DurationUnits.Week
                UndoRedo.OptionChanged(.DurationUnit)
            Else
                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "Duration", selActivityDetail.Duration)
                .Duration = intNrDays
                UndoRedo.ValueChanged(.Duration)

                UndoRedo.UndoBuffer_Initialise(selActivityDetail, "DurationUnit", selActivityDetail.DurationUnit)
                .DurationUnit = ActivityDetail.DurationUnits.Day
                UndoRedo.OptionChanged(.DurationUnit)
            End If
        End With
    End Sub

    Protected Overrides Sub OnCancelRowEdit(ByVal e As System.Windows.Forms.QuestionEventArgs)

        If Me.EditRowFlag = Me.Rows.Count - 2 AndAlso Me.EditRowFlag = Me.Grid.Count Then
            ' If the user has canceled the edit of a newly created row, 
            ' replace the corresponding logframe row with a new, empty one.
            Me.EditRow = New PlanningGridRow
        Else
            ' If the user has canceled the edit of an existing row, 
            ' release the corresponding logframe row.
            Me.EditRow = Nothing
            Me.EditRowFlag = -1
        End If
        Me.Reload()
        MyBase.OnCancelRowEdit(e)

    End Sub

    Protected Overrides Sub OnNewRowNeeded(ByVal e As System.Windows.Forms.DataGridViewRowEventArgs)
        Me.EditRow = New PlanningGridRow()
        Me.EditRowFlag = Me.Rows.Count - 1
    End Sub

    Protected Overrides Sub OnRowDirtyStateNeeded(ByVal e As System.Windows.Forms.QuestionEventArgs)
        If Not rowScopeCommit Then

            ' In cell-level commit scope, indicate whether the value
            ' of the current cell has been modified.
            e.Response = Me.IsCurrentCellDirty

        End If
    End Sub

    Protected Overrides Sub OnRowValidated(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        ' Save row changes if any were made and release the edited 
        ' planning grid row if there is one.
        If e.RowIndex >= Me.Grid.Count AndAlso e.RowIndex <> Me.Rows.Count - 1 Then

            ' Add the new planning grid row to grid.
            Me.Grid.Add(Me.EditRow)

            Me.EditRow = Nothing
            Me.EditRowFlag = -1
            Me.Reload()
        ElseIf (Me.EditRow IsNot Nothing) AndAlso e.RowIndex < Me.Grid.Count Then

            ' Save the modified planning grid row in grid.
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
