Imports System.Text.RegularExpressions

Partial Public Class DataGridViewLogframe

#Region "Mouse actions"
    Protected Overrides Sub OnCellClick(ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        MyBase.OnCellClick(e)

        If CurrentCell.IsInEditMode = False Then
            MoveCurrentCell()
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.Tab Then
            MoveCurrentCell()
            CurrentCell.Selected = True

            InvalidateSelectionRectangle()
        End If
        MyBase.OnKeyUp(e)
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim hit As DataGridView.HitTestInfo = HitTest(e.X, e.Y)

        If hit.ColumnIndex > 0 And hit.RowIndex > 0 Then
            Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
            Dim selGridRow As LogframeRow = Me.Grid(hit.RowIndex)
            Dim strStructs As String = String.Empty

            If e.Button = MouseButtons.Right Then
                If hit.Type = DataGridViewHitTestType.Cell Then
                    ' Create a rectangle using the DragSize, with the mouse position being at the center of the rectangle.
                    Dim dragSize As Size = SystemInformation.DragSize
                    DragBoxFromMouseDown = New Rectangle(New Point(e.X - (dragSize.Width / 2), _
                                                                    e.Y - (dragSize.Height / 2)), dragSize)

                    If SelectionRectangle.Rectangle.Contains(e.X, e.Y) = False Then Me.CurrentCell = selCell
                End If
            Else
                'determine where exactly in the text the user clicked
                If SelectionRectangle.Rectangle.Contains(e.X, e.Y) = True Then
                    ClickPoint = New Point(e.X, e.Y)
                Else
                    ClickPoint = Nothing
                End If
                DragBoxFromMouseDown = Rectangle.Empty
            End If

        End If
        MyBase.OnMouseDown(e)
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
        If ((e.Button And MouseButtons.Right) = MouseButtons.Right) Then
            OnMouseMove_DragItems(e.Location)
        End If
        MyBase.OnMouseMove(e)

        InvalidateSelectionRectangle()
    End Sub

    Private Sub OnMouseMove_DragItems(ByVal MouseLocation As Point)
        Dim p1 As New Point(Me.Location.X, Me.Location.Y)
        Dim p2 As New Point(Me.Location.X, Me.Location.Y + Me.Height)

        'if cursor nears the rim of the datagridview while dragging, scroll up or down
        If MouseLocation.Y < p1.Y + 50 And FirstDisplayedScrollingRowIndex > 0 Then _
            FirstDisplayedScrollingRowIndex -= 1
        If MouseLocation.Y > p2.Y - 50 And Rows(Me.RowCount - 1).Displayed = False Then _
            FirstDisplayedScrollingRowIndex += 1

        'select cursor
        If Control.ModifierKeys = Keys.Control Then
            Cursor.Current = Cursors.HSplit
            DragOperator = DragOperatorValues.Copy
        Else
            Cursor.Current = Cursors.Hand
            DragOperator = DragOperatorValues.Move
        End If

        Drag(MouseLocation)
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        DragReleased = True

        InvalidateSelectionRectangle()
        If e.Button = MouseButtons.Right Then
            If Control.ModifierKeys = Keys.Control Then
                DragOperator = DragOperatorValues.Copy
            Else
                DragOperator = DragOperatorValues.Move
            End If
            Drop(e.X, e.Y)
        End If
        DragBoxFromMouseDown = Rectangle.Empty

        MyBase.OnMouseUp(e)
    End Sub

    Protected Overrides Sub OnColumnWidthChanged(ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs)
        If Me.IsCurrentCellInEditMode = False Then
            MyBase.OnColumnWidthChanged(e)

            boolColumnWidthChanged = True
        End If
    End Sub

    Protected Overrides Sub OnScroll(ByVal e As System.Windows.Forms.ScrollEventArgs)
        MyBase.OnScroll(e)

        InvalidateSelectionRectangle()
    End Sub
#End Region

#Region "Drag and drop"
    Public Sub Drag(ByVal MouseLocation As Point)
        ' If the mouse moves outside the rectangle, start the drag.
        If (Rectangle.op_Inequality(DragBoxFromMouseDown, Rectangle.Empty) And _
            Not DragBoxFromMouseDown.Contains(MouseLocation)) Then

            If DragReleased = True Then
                If DragOperator = DragOperatorValues.Copy Then
                    CopyItems()
                Else
                    CutItems(False)
                End If
                DragReleased = False
            End If

            Dim hit As DataGridView.HitTestInfo = HitTest(MouseLocation.X, MouseLocation.Y)
            If hit.Type = DataGridViewHitTestType.Cell Then
                Dim selCell As DataGridViewCell = Me(hit.ColumnIndex, hit.RowIndex)
                SelectionRectangle.FirstRowIndex = selCell.RowIndex
                SelectionRectangle.LastRowIndex = selCell.RowIndex
                SelectionRectangle.FirstColumnIndex = selCell.ColumnIndex
                SelectionRectangle.LastColumnIndex = selCell.ColumnIndex

                If Columns(SelectionRectangle.FirstColumnIndex).Name.Contains("RTF") Then SelectionRectangle.FirstColumnIndex = SelectionRectangle.FirstColumnIndex - 1
                If Columns(SelectionRectangle.LastColumnIndex).Name.Contains("Sort") Then SelectionRectangle.LastColumnIndex = SelectionRectangle.LastColumnIndex + 1

                SetSelectionRectangleGridArea_Modify_Merged()
                If SelectionRectangle.LastRowIndex > Me.RowCount - 1 Then SelectionRectangle.LastRowIndex = Me.RowCount - 1


                'draw insert/copy indicator line
                Dim rSelStart As Rectangle = GetCellDisplayRectangle(SelectionRectangle.FirstColumnIndex, SelectionRectangle.FirstRowIndex, False)
                Dim rSelEnd As Rectangle = GetCellDisplayRectangle(SelectionRectangle.LastColumnIndex, SelectionRectangle.LastRowIndex, False)

                Dim intVertDivider As Integer = rSelStart.Top + ((rSelEnd.Bottom - rSelStart.Top) / 2)
                pStartInsertLine.X = rSelStart.Left
                pEndInsertLine.X = rSelEnd.Right - 1
                If MouseLocation.Y <= intVertDivider Then pStartInsertLine.Y = rSelStart.Top Else pStartInsertLine.Y = rSelEnd.Bottom
                pEndInsertLine.Y = pStartInsertLine.Y

                Dim graph As Graphics = CreateGraphics()
                If Not (pStartInsertLine = Nothing Or pEndInsertLine = Nothing) Then
                    If pStartInsertLine <> pStartInsertLineOld Then
                        InvalidateColumn(SelectionRectangle.FirstColumnIndex)
                        InvalidateColumn(SelectionRectangle.LastColumnIndex)
                        Update()

                        Dim penGreen2 As New Pen(Color.Green, 2)
                        graph.DrawLine(penGreen2, pStartInsertLine, pEndInsertLine)
                        Dim triangleStart(2) As Point
                        triangleStart(0) = New Point(pStartInsertLine.X, pStartInsertLine.Y - 6)
                        triangleStart(1) = New Point(pStartInsertLine.X, pStartInsertLine.Y + 6)
                        triangleStart(2) = New Point(pStartInsertLine.X + 6, pStartInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleStart)
                        Dim triangleEnd(2) As Point
                        triangleEnd(0) = New Point(pEndInsertLine.X, pEndInsertLine.Y - 6)
                        triangleEnd(1) = New Point(pEndInsertLine.X, pEndInsertLine.Y + 6)
                        triangleEnd(2) = New Point(pEndInsertLine.X - 8, pEndInsertLine.Y)
                        graph.FillPolygon(Brushes.Green, triangleEnd)
                        pStartInsertLineOld = pStartInsertLine
                        pEndInsertLineOld = pEndInsertLine
                    End If
                End If

            Else
                pStartInsertLine = Nothing
                pEndInsertLine = Nothing
            End If
        End If
    End Sub

    Public Sub Drop(ByVal mouseX As Integer, ByVal mouseY As Integer)
        Dim hit As DataGridView.HitTestInfo = HitTest(pStartInsertLine.X, pStartInsertLine.Y)

        If hit.Type = DataGridViewHitTestType.Cell Then
            Dim CopyGroup As ClipboardItems = ItemClipboard.GetCopyGroup()
            Dim selCell As DataGridViewCell

            If hit.RowIndex <= Me.RowCount - 1 Then
                selCell = Me(hit.ColumnIndex, hit.RowIndex)
            Else
                selCell = Me(hit.ColumnIndex, Me.RowCount - 1)
            End If

            PasteItems(CopyGroup, ClipboardItems.PasteOptions.PasteAll, selCell)
        End If
    End Sub
#End Region

#Region "General methods"
    Public Function IsResourceBudgetRow(ByVal selRow As LogframeRow) As Boolean
        If selRow.Section = Logframe.SectionTypes.ActivitiesSection And Me.ShowResourcesBudget = True Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SetFocusOnItem(ByVal selItem As LogframeObject, Optional ByVal FocusOnBudgetTotal As Boolean = False)
        Dim selGridRow As LogframeRow
        Dim intColIndex As Integer, intRowIndex As Integer
        For intRowIndex = 0 To Grid.Count - 1
            selGridRow = Grid(intRowIndex)

            Select Case selItem.GetType
                Case GetType(Goal), GetType(Purpose), GetType(Output), GetType(Activity)
                    If selGridRow.Struct Is selItem Then
                        intColIndex = 1
                        Exit For
                    End If
                Case GetType(Indicator)
                    If selGridRow.Indicator Is selItem Then
                        intColIndex = 3
                        Exit For
                    End If
                Case GetType(Resource)
                    If selGridRow.Resource Is selItem Then
                        If FocusOnBudgetTotal = False Then
                            intColIndex = 3
                        Else
                            intColIndex = 5
                        End If

                        Exit For
                    End If
                Case GetType(VerificationSource)
                        If selGridRow.VerificationSource Is selItem Then
                            intColIndex = 5
                            Exit For
                        End If
                Case GetType(Assumption)
                        If selGridRow.Assumption Is selItem Then
                            intColIndex = 7
                            Exit For
                        End If
            End Select
        Next

        If intRowIndex >= 0 And intRowIndex < RowCount Then
            EnsureColumnIsVisible(intColIndex)
            CurrentCell = Me(intColIndex, intRowIndex)
            MoveCurrentCell()
        End If
    End Sub

    Public Function GetFirstItem() As LogframeObject
        For Each selGridRow As LogframeRow In Me.Grid
            If selGridRow.RowType = LogframeRow.RowTypes.Normal Then
                If String.IsNullOrEmpty(selGridRow.StructRtf) = False Then
                    Return selGridRow.Struct
                End If
                If IsResourceBudgetRow(selGridRow) = False Then
                    If String.IsNullOrEmpty(selGridRow.IndicatorRtf) = False Then
                        Return selGridRow.Indicator
                    ElseIf String.IsNullOrEmpty(selGridRow.VerificationSourceRtf) = False Then
                        Return selGridRow.VerificationSource
                    End If
                Else
                    If String.IsNullOrEmpty(selGridRow.ResourceRtf) = False Then
                        Return selGridRow.Resource
                    End If
                End If
                If String.IsNullOrEmpty(selGridRow.AssumptionRtf) = False Then
                    Return selGridRow.Assumption
                End If
            End If
        Next

        Return Nothing
    End Function

    Protected Overloads Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
        If keyData = Keys.Enter Then

            If CurrentRow.Index = RowCount - 1 Then
                Dim selGridRow = Me.Grid(CurrentRow.Index)
                Dim intRowIndex As Integer = CurrentCell.RowIndex, intColIndex As Integer = CurrentCell.ColumnIndex
                CurrentCell = Me(intColIndex, intRowIndex - 1)
                CurrentCell = Me(intColIndex, intRowIndex)
            End If
            MyBase.ProcessEnterKey(Keys.Enter)

            Return True
        End If

        Return MyBase.ProcessDialogKey(keyData)
    End Function

    Protected Overloads Overrides Function ProcessDataGridViewKey(ByVal e As KeyEventArgs) As Boolean

        If e.KeyCode = Keys.Enter Then
            MyBase.ProcessEnterKey(Keys.Enter)
            CurrentCell.Selected = True

            InvalidateSelectionRectangle()
            Return True
        End If

        Return MyBase.ProcessDataGridViewKey(e)
    End Function

    Public Sub EnsureColumnIsVisible(ByVal intColumnIndex As Integer)
        Select Case intColumnIndex
            Case 3
                If ShowIndicatorColumn = False Then
                    ShowIndicatorColumn = True
                    Reload()
                    RaiseEvent ShowIndicatorColumnChanged()
                End If
            Case 5
                If ShowVerificationSourceColumn = False Then
                    ShowVerificationSourceColumn = True
                    Reload()
                    RaiseEvent ShowVerificationSourceColumnChanged()
                End If
            Case 7
                If ShowAssumptionColumn = False Then
                    ShowAssumptionColumn = True
                    Reload()
                    RaiseEvent ShowAssumptionColumnChanged()
                End If
        End Select
    End Sub

    Public Sub EnsureSectionIsVisible(ByVal intSection As Integer)
        Select Case intSection
            Case Logframe.SectionTypes.GoalsSection
                If ShowGoals = False Then ShowGoals = True
                Reload()
            Case Logframe.SectionTypes.PurposesSection
                If ShowPurposes = False Then ShowPurposes = True
                Reload()
            Case Logframe.SectionTypes.OutputsSection
                If ShowOutputs = False Then ShowOutputs = True
                Reload()
            Case Logframe.SectionTypes.ActivitiesSection
                If ShowActivities = False Then ShowActivities = True
                Reload()

        End Select
    End Sub

    Public Sub EnsureIndicatorsResourcesAreVisible(ByVal objItem As LogframeObject)
        Select Case objItem.GetType
            Case GetType(Indicator), GetType(VerificationSource)
                If ShowResourcesBudget = True Then
                    ShowResourcesBudget = False
                    Reload()
                    RaiseEvent ShowResourcesBudgetChanged()
                End If
            Case GetType(Resource)
                If ShowResourcesBudget = False Then
                    ShowResourcesBudget = True
                    Reload()
                    RaiseEvent ShowResourcesBudgetChanged()
                End If
        End Select
    End Sub
#End Region

#Region "Add and insert items"
    Public Sub AddNewItem()
        If CurrentCell Is Nothing Then Exit Sub

        Dim intRowIndex As Integer
        Dim intColIndex = CurrentCell.ColumnIndex
        Dim strColName As String = Me.Columns(intColIndex).Name
        Dim selGridrow As LogframeRow

        For intRowIndex = CurrentCell.RowIndex To RowCount - 1
            selGridrow = Grid(intRowIndex)
            Select Case strColName
                Case "StructSort", "StructRTF"
                    If selGridrow.Struct Is Nothing Then Exit For
                Case "IndSort", "IndRTF"
                    If IsResourceBudgetRow(selGridrow) = False Then
                        If selGridrow.Indicator Is Nothing Then Exit For
                    Else
                        If selGridrow.Resource Is Nothing Then Exit For
                    End If
                Case "VerSort", "VerRTF"
                    If IsResourceBudgetRow(selGridrow) = False Then
                        If selGridrow.VerificationSource Is Nothing Then Exit For
                    Else
                        If selGridrow.Resource Is Nothing Then Exit For
                    End If
                Case "AsmSort", "AsmRTF"
                    If selGridrow.Assumption Is Nothing Then Exit For
            End Select
        Next

        CurrentCell = Me(intColIndex, intRowIndex)
        Me.BeginEdit(False)
    End Sub

    Public Sub InsertItem()
        Dim selObject As LogframeObject = Nothing
        Dim boolBudget As Boolean
        Dim intIndex As Integer

        If CurrentItem(False).GetType = GetType(Single) Then
            'budget item
            Dim selGridRow As LogframeRow = Me.Grid(CurrentRow.Index)
            If selGridRow.Resource IsNot Nothing Then selObject = selGridRow.Resource
            boolBudget = True
        Else
            selObject = CType(Me.CurrentItem(False), LogframeObject)
        End If

        If selObject Is Nothing Then Exit Sub

        Select Case selObject.GetType
            Case GetType(Goal)
                Dim selGoal As Goal = DirectCast(selObject, Goal)
                intIndex = Me.Logframe.Goals.IndexOf(selGoal)

                Dim NewGoal As New Goal
                Me.Logframe.Goals.Insert(intIndex, NewGoal)
                UndoRedo.ItemInserted(NewGoal, Me.Logframe.Goals)
                selObject = NewGoal
            Case GetType(Purpose)
                Dim selPurpose As Purpose = DirectCast(selObject, Purpose)
                intIndex = Me.Logframe.Purposes.IndexOf(selPurpose)

                Dim NewPurpose As New Purpose
                Me.Logframe.Purposes.Insert(intIndex, NewPurpose)
                UndoRedo.ItemInserted(NewPurpose, Me.Logframe.Purposes)
                selObject = NewPurpose
            Case GetType(Output)
                Dim selOutput As Output = DirectCast(selObject, Output)
                Dim ParentPurpose As Purpose = Me.Logframe.GetPurposeByGuid(selOutput.ParentPurposeGuid)
                If ParentPurpose IsNot Nothing Then
                    intIndex = ParentPurpose.Outputs.IndexOf(selOutput)

                    Dim NewOutput As New Output
                    ParentPurpose.Outputs.Insert(intIndex, NewOutput)
                    UndoRedo.ItemInserted(NewOutput, ParentPurpose.Outputs)
                    selObject = NewOutput
                End If
            Case GetType(Activity)
                Dim selActivity As Activity = DirectCast(selObject, Activity)
                If selActivity.ParentOutputGuid <> Guid.Empty Then
                    Dim ParentOutput As Output = Me.Logframe.GetOutputByGuid(selActivity.ParentOutputGuid)
                    If ParentOutput IsNot Nothing Then
                        intIndex = ParentOutput.Activities.IndexOf(selActivity)

                        Dim NewActivity As New Activity
                        ParentOutput.Activities.Insert(intIndex, NewActivity)
                        UndoRedo.ItemInserted(NewActivity, ParentOutput.Activities)
                        selObject = NewActivity
                    End If
                Else
                    Dim ParentActivity As Activity = Me.Logframe.GetActivityByGuid(selActivity.ParentActivityGuid)
                    If ParentActivity IsNot Nothing Then
                        intIndex = ParentActivity.Activities.IndexOf(selActivity)

                        Dim NewActivity As New Activity
                        ParentActivity.Activities.Insert(intIndex, NewActivity)
                        UndoRedo.ItemInserted(NewActivity, ParentActivity.Activities)
                        selObject = NewActivity
                    End If
                End If
            Case GetType(Indicator)
                Dim selIndicator As Indicator = DirectCast(selObject, Indicator)
                If selIndicator.ParentStructGuid <> Guid.Empty Then
                    Dim ParentStruct As Struct = Me.Logframe.GetStructByGuid(selIndicator.ParentStructGuid)
                    If ParentStruct IsNot Nothing Then
                        intIndex = ParentStruct.Indicators.IndexOf(selIndicator)

                        Dim NewIndicator As New Indicator
                        ParentStruct.Indicators.Insert(intIndex, NewIndicator)
                        UndoRedo.ItemInserted(NewIndicator, ParentStruct.Indicators)
                        selObject = NewIndicator
                    End If
                Else
                    Dim ParentIndicator As Indicator = Me.Logframe.GetIndicatorByGuid(selIndicator.ParentIndicatorGuid)
                    If ParentIndicator IsNot Nothing Then
                        intIndex = ParentIndicator.Indicators.IndexOf(selIndicator)

                        Dim NewIndicator As New Indicator
                        ParentIndicator.Indicators.Insert(intIndex, NewIndicator)
                        UndoRedo.ItemInserted(NewIndicator, ParentIndicator.Indicators)
                        selObject = NewIndicator
                    End If
                End If
            Case GetType(Resource)
                Dim selResource As Resource = DirectCast(selObject, Resource)
                Dim ParentActivity As Activity = Me.Logframe.GetActivityByGuid(selResource.ParentStructGuid)
                If ParentActivity IsNot Nothing Then
                    intIndex = ParentActivity.Resources.IndexOf(selResource)

                    Dim NewResource As New Resource
                    ParentActivity.Resources.Insert(intIndex, NewResource)
                    UndoRedo.ItemInserted(NewResource, ParentActivity.Resources)
                    selObject = NewResource
                End If
            Case GetType(VerificationSource)
                Dim selVerificationSource As VerificationSource = DirectCast(selObject, VerificationSource)
                Dim ParentIndicator As Indicator = Me.Logframe.GetIndicatorByGuid(selVerificationSource.ParentIndicatorGuid)
                If ParentIndicator IsNot Nothing Then
                    intIndex = ParentIndicator.VerificationSources.IndexOf(selVerificationSource)

                    Dim NewVerificationSource As New VerificationSource
                    ParentIndicator.VerificationSources.Insert(intIndex, NewVerificationSource)
                    UndoRedo.ItemInserted(NewVerificationSource, ParentIndicator.VerificationSources)
                    selObject = NewVerificationSource
                End If
            Case GetType(Assumption)
                Dim selAssumption As Assumption = DirectCast(selObject, Assumption)
                Dim ParentStruct As Struct = Me.Logframe.GetStructByGuid(selAssumption.ParentStructGuid)
                If ParentStruct IsNot Nothing Then
                    intIndex = ParentStruct.Assumptions.IndexOf(selAssumption)

                    Dim NewAssumption As New Assumption
                    ParentStruct.Assumptions.Insert(intIndex, NewAssumption)
                    UndoRedo.ItemInserted(NewAssumption, ParentStruct.Assumptions)
                    selObject = NewAssumption
                End If
        End Select

        Reload()

        SetFocusOnItem(selObject)
        If boolBudget = True Then
            Me.CurrentCell = Me(5, CurrentCell.RowIndex)
        End If

        Me.BeginEdit(False)
    End Sub

    Public Sub InsertParentItem()
        If CurrentItem(False) IsNot Nothing Then
            Select Case CurrentItem(False).GetType
                Case GetType(Activity)
                    Dim selActivity As Activity
                    Dim ParentActivities As Activities = Nothing

                    selActivity = TryCast(Me.CurrentItem(False), Activity)

                    If selActivity IsNot Nothing Then
                        ParentActivities = CurrentLogFrame.GetParentCollection(selActivity)

                        If ParentActivities IsNot Nothing Then
                            Dim intIndex As Integer = ParentActivities.IndexOf(selActivity)
                            Dim NewParent As New Activity

                            ParentActivities.Remove(selActivity)

                            ParentActivities.Insert(intIndex, NewParent)
                            UndoRedo.ItemInserted(NewParent, ParentActivities)

                            NewParent.Activities.Add(selActivity)
                            UndoRedo.ItemParentInserted(selActivity, ParentActivities, intIndex, NewParent.Activities)

                            Reload()
                            SetFocusOnItem(NewParent)
                            Me.BeginEdit(False)
                        End If
                    End If
                Case GetType(Indicator)
                    Dim selIndicator As Indicator
                    Dim ParentIndicators As Indicators = Nothing

                    selIndicator = TryCast(Me.CurrentItem(False), Indicator)

                    If selIndicator IsNot Nothing Then
                        ParentIndicators = CurrentLogFrame.GetParentCollection(selIndicator)

                        If ParentIndicators IsNot Nothing Then
                            Dim intIndex As Integer = ParentIndicators.IndexOf(selIndicator)
                            Dim NewParent As New Indicator

                            CopySettingsOfRelatedIndicator(selIndicator, NewParent)

                            ParentIndicators.Remove(selIndicator)

                            ParentIndicators.Insert(intIndex, NewParent)
                            UndoRedo.ItemInserted(NewParent, ParentIndicators)

                            NewParent.Indicators.Add(selIndicator)
                            UndoRedo.ItemParentInserted(selIndicator, ParentIndicators, intIndex, NewParent.Indicators)

                            Reload()
                            SetFocusOnItem(NewParent)
                            Me.BeginEdit(False)
                        End If
                    End If
            End Select
        End If
    End Sub

    Public Sub InsertChildItem()
        If CurrentItem(False) IsNot Nothing Then
            Select Case CurrentItem(False).GetType
                Case GetType(Activity)
                    Dim selActivity As Activity

                    selActivity = TryCast(Me.CurrentItem(False), Activity)

                    If selActivity IsNot Nothing Then
                        Dim ChildActivity As New Activity

                        selActivity.Activities.Insert(0, ChildActivity)
                        UndoRedo.ItemChildInserted(ChildActivity, selActivity.Activities)

                        Reload()
                        SetFocusOnItem(ChildActivity)
                        Me.BeginEdit(False)
                    End If
                Case GetType(Indicator)
                    Dim selIndicator As Indicator

                    selIndicator = TryCast(Me.CurrentItem(False), Indicator)

                    If selIndicator IsNot Nothing Then
                        Dim ChildIndicator As New Indicator

                        CopySettingsOfRelatedIndicator(selIndicator, ChildIndicator)

                        selIndicator.Indicators.Insert(0, ChildIndicator)
                        UndoRedo.ItemChildInserted(ChildIndicator, selIndicator.Indicators)

                        Reload()
                        SetFocusOnItem(ChildIndicator)
                        Me.BeginEdit(False)
                    End If
            End Select
        End If

    End Sub

    Private Sub CopySettingsOfRelatedIndicator(ByVal objSourceIndicator As Indicator, ByVal objNewIndicator As Indicator)
        'settings for parent indicator and question type of sub-indicator
        Select Case objSourceIndicator.QuestionType
            Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                If objSourceIndicator.ScoringSystem <> Indicator.ScoringSystems.Score Then
                    objNewIndicator.QuestionType = objSourceIndicator.QuestionType

                End If
            Case Else
                objNewIndicator.QuestionType = objSourceIndicator.QuestionType
        End Select

        With objNewIndicator
            .idParent = objSourceIndicator.idIndicator
            .ParentIndicatorGuid = objSourceIndicator.Guid
            .Section = objSourceIndicator.Section
            .ScoringSystem = objSourceIndicator.ScoringSystem
            .Registration = objSourceIndicator.Registration
            .TargetGroupGuid = objSourceIndicator.TargetGroupGuid
            .AggregateVertical = objSourceIndicator.AggregateVertical

            'type specific settings
            Select Case objSourceIndicator.QuestionType
                Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue
                    If .ValuesDetail Is Nothing Then .ValuesDetail = New ValuesDetail

                    .ValuesDetail.NrDecimals = objSourceIndicator.ValuesDetail.NrDecimals
                    .ValuesDetail.Unit = objSourceIndicator.ValuesDetail.Unit
                    .ValuesDetail.ValueName = objSourceIndicator.ValuesDetail.ValueName
                Case Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula
                    Using copier As New ObjectCopy

                        .ValuesDetail = copier.CopyObject(objSourceIndicator.ValuesDetail)

                        .Statements.Clear()
                        .Statements = copier.CopyCollection(objSourceIndicator.Statements)
                    End Using
                Case Indicator.QuestionTypes.Scale, Indicator.QuestionTypes.LikertScale, Indicator.QuestionTypes.CumulativeScale, Indicator.QuestionTypes.SemanticDiff
                    Using copier As New ObjectCopy
                        .ValuesDetail = copier.CopyObject(objSourceIndicator.ValuesDetail)
                        .ScalesDetail = copier.CopyObject(objSourceIndicator.ScalesDetail)

                        .Statements.Clear()
                        .Statements = copier.CopyCollection(objSourceIndicator.Statements)
                        .ResponseClasses.Clear()
                        .ResponseClasses = copier.CopyCollection(objSourceIndicator.ResponseClasses)
                    End Using
                Case Indicator.QuestionTypes.MixedSubIndicators
                    'do nothing
                Case Else
                    If .ResponseClasses.Count = 0 Then
                        For i = 0 To objSourceIndicator.ResponseClasses.Count - 1
                            .ResponseClasses.Add(New ResponseClass(objSourceIndicator.ResponseClasses(i).ClassName, objSourceIndicator.ResponseClasses(i).Value))
                        Next
                    End If
            End Select
        End With
    End Sub
#End Region

#Region "Move items"
    Public Sub MoveItem(ByVal intDirection As Integer)
        Dim selGridrow As LogframeRow = Grid(CurrentCell.RowIndex)
        Dim selItem As LogframeObject = TryCast(Me.CurrentItem(False), LogframeObject)
        If selItem Is Nothing Then selItem = selGridrow.Resource
        If selItem Is Nothing Then Exit Sub

        Select Case selItem.GetType
            Case GetType(Goal)
                If intDirection < 0 Then
                    selItem = MoveGoal_ToPreviousParent(selItem)
                Else
                    selItem = MoveGoal_ToNextParent(selItem)
                End If
            Case GetType(Purpose)
                If intDirection < 0 Then
                    selItem = MovePurpose_ToPreviousParent(selItem)
                Else
                    selItem = MovePurpose_ToNextParent(selItem)
                End If
            Case GetType(Output)
                If intDirection < 0 Then
                    selItem = MoveOutput_ToPreviousParent(selItem)
                Else
                    selItem = MoveOutput_ToNextParent(selItem)
                End If
            Case GetType(Activity)
                If intDirection < 0 Then
                    selItem = MoveActivity_ToPreviousParent(selItem)
                Else
                    selItem = MoveActivity_ToNextParent(selItem)
                End If
            Case GetType(Indicator)
                If intDirection < 0 Then
                    selItem = MoveIndicator_ToPreviousParent(selItem)
                Else
                    selItem = MoveIndicator_ToNextParent(selItem)
                End If
            Case Else
                Dim objParentCollection As Object = CurrentLogFrame.GetParentCollection(selItem)
                Dim intIndex, intOldIndex As Integer

                If objParentCollection IsNot Nothing Then
                    intIndex = objParentCollection.IndexOf(selItem)
                    intOldIndex = intIndex

                    If intDirection < 0 And intIndex = 0 Then
                        MoveItem_ToPreviousParent(selItem)
                        Exit Sub
                    ElseIf intDirection > 0 And intIndex = objParentCollection.Count - 1 Then
                        MoveItem_ToNextParent(selItem)
                    Else
                        objParentCollection.RemoveAt(intIndex)
                        intIndex += intDirection
                        objParentCollection.Insert(intIndex, selItem)

                        If intDirection < 0 Then
                            UndoRedo.ItemMovedUp(selItem, selItem, objParentCollection, intOldIndex, objParentCollection)
                        Else
                            UndoRedo.ItemMovedDown(selItem, selItem, objParentCollection, intOldIndex, objParentCollection)
                        End If
                    End If
                End If
        End Select

        Reload()
        SetFocusOnItem(selItem)
    End Sub

    Private Sub MoveItem_ToPreviousParent(ByVal selItem As LogframeObject)
        Select Case selItem.GetType
            Case GetType(Indicator)
                MoveIndicator_ToPreviousParent(selItem)
            Case GetType(Resource)
                selItem = MoveResource_ToPreviousParent(selItem)
            Case GetType(VerificationSource)
                MoveVerificationSource_ToPreviousParent(selItem)
            Case GetType(Assumption)
                MoveAssumption_ToPreviousParent(selItem)
        End Select

        Reload()
        SetFocusOnItem(selItem)
    End Sub

    Private Sub MoveItem_ToNextParent(ByVal selItem As LogframeObject)
        Select Case selItem.GetType
            Case GetType(Indicator)
                selItem = MoveIndicator_ToNextParent(selItem)
            Case GetType(Resource)
                MoveResource_ToNextParent(selItem)
            Case GetType(VerificationSource)
                MoveVerificationSource_ToNextParent(selItem)
            Case GetType(Assumption)
                MoveAssumption_ToNextParent(selItem)
        End Select

        Reload()
        SetFocusOnItem(selItem)
    End Sub

    Private Function MoveGoal_ToPreviousParent(ByVal selGoal As Goal) As LogframeObject
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)
        Dim intIndex As Integer

        Select Case PreviousStruct.GetType
            Case GetType(Goal)
                intIndex = Me.Logframe.Goals.IndexOf(PreviousStruct)
                If intIndex < 0 Then
                    'section title row

                Else
                    Dim intOldIndex As Integer = Me.Logframe.Goals.IndexOf(selGoal)
                    Me.Logframe.Goals.Remove(selGoal)
                    Me.Logframe.Goals.Insert(intIndex, selGoal)

                    UndoRedo.ItemMovedUp(selGoal, selGoal, Me.Logframe.Goals, intOldIndex, Me.Logframe.Goals)
                End If
        End Select

        Return selGoal
    End Function

    Private Function MoveGoal_ToNextParent(ByVal selGoal As Goal) As LogframeObject
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim NextStruct As Struct = Me.Grid.GetNextStruct(intRowIndex)
        Dim PurposeSectionRow As LogframeRow = Me.Grid(Me.Grid.GetRowIndexOfSectionTitle(Logframe.SectionTypes.PurposesSection))
        Dim intIndex As Integer

        Select Case NextStruct.GetType
            Case GetType(Purpose)
                If CType(NextStruct, Purpose) Is PurposeSectionRow.Struct Then
                    'section title row
                    intRowIndex = Me.Grid.GetRowIndexOfFirstRowInSection(Logframe.SectionTypes.PurposesSection)
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveGoal_ToNextParent(selGoal)
                Else
                    Dim NewPurpose As New Purpose(selGoal.RTF)
                    Dim intOldIndex As Integer = Me.Logframe.Goals.IndexOf(selGoal)

                    Using copier As New ObjectCopy
                        NewPurpose.Indicators = copier.CopyCollection(selGoal.Indicators)
                        NewPurpose.Assumptions = copier.CopyCollection(selGoal.Assumptions)
                    End Using

                    Me.Logframe.Purposes.Insert(0, NewPurpose)
                    Me.Logframe.Goals.Remove(selGoal)

                    UndoRedo.ItemMovedDown(selGoal, NewPurpose, Me.Logframe.Goals, intOldIndex, Me.Logframe.Purposes)

                    Return NewPurpose
                End If
            Case GetType(Goal)
                Dim intOldIndex As Integer = Me.Logframe.Goals.IndexOf(selGoal)

                intIndex = Me.Logframe.Goals.IndexOf(NextStruct)
                Me.Logframe.Goals.Remove(selGoal)
                Me.Logframe.Goals.Insert(intIndex, selGoal)

                UndoRedo.ItemMovedDown(selGoal, selGoal, Me.Logframe.Goals, intOldIndex, Me.Logframe.Goals)
        End Select

        Return selGoal
    End Function

    Private Function MovePurpose_ToPreviousParent(ByVal selPurpose As Purpose) As LogframeObject
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)
        Dim intIndex As Integer

        Select Case PreviousStruct.GetType
            Case GetType(Purpose)
                intIndex = Me.Logframe.Purposes.IndexOf(PreviousStruct)
                If intIndex < 0 Then
                    'section title row
                    intRowIndex -= 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MovePurpose_ToPreviousParent(selPurpose)
                Else
                    Dim intOldIndex As Integer = Me.Logframe.Purposes.IndexOf(selPurpose)

                    Me.Logframe.Purposes.Remove(selPurpose)
                    Me.Logframe.Purposes.Insert(intIndex, selPurpose)

                    UndoRedo.ItemMovedUp(selPurpose, selPurpose, Me.Logframe.Purposes, intOldIndex, Me.Logframe.Purposes)
                End If
            Case GetType(Goal)
                Dim intOldIndex As Integer = Me.Logframe.Purposes.IndexOf(selPurpose)
                Dim NewGoal As New Goal(selPurpose.RTF)
                Using copier As New ObjectCopy
                    NewGoal.Indicators = copier.CopyCollection(selPurpose.Indicators)
                    NewGoal.Assumptions = copier.CopyCollection(selPurpose.Assumptions)
                End Using

                If selPurpose.Outputs.Count = 0 Then
                    Me.Logframe.Purposes.Remove(selPurpose)
                Else
                    selPurpose.RTF = String.Empty
                    selPurpose.Indicators.Clear()
                    selPurpose.Assumptions.Clear()
                End If

                Me.Logframe.Goals.Add(NewGoal)

                UndoRedo.ItemMovedUp(selPurpose, NewGoal, Me.Logframe.Purposes, intOldIndex, Me.Logframe.Goals)

                Return NewGoal
        End Select

        Return selPurpose
    End Function

    Private Function MovePurpose_ToNextParent(ByVal selPurpose As Purpose) As LogframeObject
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim NextStruct As Struct = Me.Grid.GetNextStruct(intRowIndex)
        Dim intIndex As Integer

        Select Case NextStruct.GetType
            Case GetType(Output)
                If CType(NextStruct, Output).ParentPurposeGuid = Guid.Empty Then
                    'section title row
                    intRowIndex = Me.Grid.GetRowIndexOfFirstRowInSection(Logframe.SectionTypes.OutputsSection)
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MovePurpose_ToNextParent(selPurpose)
                Else
                    Dim intOldIndex As Integer = Me.Logframe.Purposes.IndexOf(selPurpose)
                    Dim NewOutput As New Output(selPurpose.RTF)

                    Using copier As New ObjectCopy
                        NewOutput.Indicators = copier.CopyCollection(selPurpose.Indicators)
                        NewOutput.Assumptions = copier.CopyCollection(selPurpose.Assumptions)
                    End Using

                    Dim FirstPurpose As Purpose = Me.Logframe.Purposes(0)
                    FirstPurpose.Outputs.Insert(0, NewOutput)

                    If FirstPurpose IsNot selPurpose And selPurpose.Outputs.Count = 0 Then
                        Me.Logframe.Purposes.Remove(selPurpose)
                    Else
                        selPurpose.Indicators.Clear()
                        selPurpose.Assumptions.Clear()
                        selPurpose.SetText(String.Empty)
                    End If
                    UndoRedo.ItemMovedDown(selPurpose, NewOutput, Me.Logframe.Purposes, intOldIndex, FirstPurpose.Outputs)
                    Return NewOutput
                End If
            Case GetType(Purpose)
                Dim intOldIndex As Integer = Me.Logframe.Purposes.IndexOf(selPurpose)

                intIndex = Me.Logframe.Purposes.IndexOf(NextStruct)
                Me.Logframe.Purposes.Remove(selPurpose)
                Me.Logframe.Purposes.Insert(intIndex, selPurpose)

                UndoRedo.ItemMovedDown(selPurpose, selPurpose, Me.Logframe.Purposes, intOldIndex, Me.Logframe.Purposes)
        End Select

        Return selPurpose
    End Function

    Private Function MoveOutput_ToPreviousParent(ByVal selOutput As Output) As LogframeObject
        Dim objOutputs As Outputs = Me.Logframe.GetParentCollection(selOutput)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim intSection As Integer = selGridRow.Section
        Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)
        Dim intIndex As Integer

        Select Case PreviousStruct.GetType
            Case GetType(Output)
                If CType(PreviousStruct, Output).ParentPurposeGuid = Guid.Empty Then
                    'section title row
                    intRowIndex -= 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveOutput_ToPreviousParent(selOutput)
                Else
                    Dim intOldIndex As Integer = objOutputs.IndexOf(selOutput)
                    Dim objPreviousOutputs As Outputs = Me.Logframe.GetParentCollection(PreviousStruct)

                    intIndex = objPreviousOutputs.IndexOf(PreviousStruct)
                    If intIndex = objPreviousOutputs.Count - 1 Then intIndex += 1
                    objOutputs.Remove(selOutput)
                    objPreviousOutputs.Insert(intIndex, selOutput)

                    UndoRedo.ItemMovedUp(selOutput, selOutput, objOutputs, intOldIndex, objPreviousOutputs)
                End If
            Case GetType(Purpose)
                If intSection = Logframe.SectionTypes.OutputsSection Then
                    intRowIndex -= 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveOutput_ToPreviousParent(selOutput)
                Else
                    Dim intOldIndex As Integer = objOutputs.IndexOf(selOutput)
                    Dim NewPurpose As New Purpose(selOutput.RTF)

                    Using copier As New ObjectCopy
                        NewPurpose.Indicators = copier.CopyCollection(selOutput.Indicators)
                        NewPurpose.Assumptions = copier.CopyCollection(selOutput.Assumptions)
                    End Using

                    If selOutput.Activities.Count = 0 Then
                        objOutputs.Remove(selOutput)
                    Else
                        selOutput.RTF = String.Empty
                        selOutput.Indicators.Clear()
                        selOutput.Assumptions.Clear()
                    End If

                    Me.Logframe.Purposes.Add(NewPurpose)

                    UndoRedo.ItemMovedUp(selOutput, NewPurpose, objOutputs, intOldIndex, Me.Logframe.Purposes)

                    Return NewPurpose
                End If
        End Select

        Return selOutput
    End Function

    Private Function MoveOutput_ToNextParent(ByVal selOutput As Output) As LogframeObject
        Dim objOutputs As Outputs = Me.Logframe.GetParentCollection(selOutput)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim intSection As Integer = selGridRow.Section
        Dim NextStruct As Struct = Me.Grid.GetNextStruct(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objOutputs.IndexOf(selOutput)

        Select Case NextStruct.GetType
            Case GetType(Activity)
                If CType(NextStruct, Activity).ParentOutputGuid = Guid.Empty And CType(NextStruct, Activity).ParentActivityGuid = Guid.Empty Then
                    'section title row
                    intRowIndex = Me.Grid.GetRowIndexOfFirstRowInSection(Logframe.SectionTypes.ActivitiesSection)
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveOutput_ToNextParent(selOutput)
                Else

                    Dim NewActivity As New Activity(selOutput.RTF)

                    Using copier As New ObjectCopy
                        NewActivity.Indicators = copier.CopyCollection(selOutput.Indicators)
                        NewActivity.Assumptions = copier.CopyCollection(selOutput.Assumptions)
                    End Using
                    Dim FirstOutput As Output = Me.Logframe.Purposes(0).Outputs(0)
                    FirstOutput.Activities.Insert(0, NewActivity)

                    If FirstOutput IsNot selOutput And selOutput.Activities.Count = 0 Then
                        objOutputs.Remove(selOutput)
                    Else
                        selOutput.Indicators.Clear()
                        selOutput.Assumptions.Clear()
                        selOutput.SetText(String.Empty)
                    End If
                    UndoRedo.ItemMovedDown(selOutput, NewActivity, objOutputs, intOldIndex, FirstOutput.Activities)

                    Return NewActivity
                End If
            Case GetType(Output)
                Dim objNextOutputs As Outputs = Me.Logframe.GetParentCollection(NextStruct)
                intIndex = objNextOutputs.IndexOf(NextStruct)

                objOutputs.Remove(selOutput)
                objNextOutputs.Insert(intIndex, selOutput)

                UndoRedo.ItemMovedDown(selOutput, selOutput, objOutputs, intOldIndex, objNextOutputs)
            Case GetType(Purpose)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveOutput_ToNextParent(selOutput)

        End Select

        Return selOutput
    End Function

    Private Function MoveActivity_ToPreviousParent(ByVal selActivity As Activity) As LogframeObject
        Dim objActivities As Activities = Me.Logframe.GetParentCollection(selActivity)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim intSection As Integer = selGridRow.Section
        Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objActivities.IndexOf(selActivity)

        Select Case PreviousStruct.GetType
            Case GetType(Activity)
                If CType(PreviousStruct, Activity).ParentOutputGuid = Guid.Empty And CType(PreviousStruct, Activity).ParentActivityGuid = Guid.Empty Then
                    'section title row
                    intRowIndex -= 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveActivity_ToPreviousParent(selActivity)
                Else
                    Dim objPreviousActivities As Activities = Me.Logframe.GetParentCollection(PreviousStruct)

                    intIndex = objPreviousActivities.IndexOf(PreviousStruct)
                    If intIndex = objPreviousActivities.Count - 1 Then intIndex += 1

                    objActivities.Remove(selActivity)
                    objPreviousActivities.Insert(intIndex, selActivity)

                    UndoRedo.ItemMovedUp(selActivity, selActivity, objActivities, intOldIndex, objPreviousActivities)
                End If
            Case GetType(Output)
                If intSection = Logframe.SectionTypes.ActivitiesSection Then
                    intRowIndex -= 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveActivity_ToPreviousParent(selActivity)
                Else
                    Dim NewOutput As New Output(selActivity.RTF)
                    Using copier As New ObjectCopy
                        NewOutput.Indicators = copier.CopyCollection(selActivity.Indicators)
                        NewOutput.Assumptions = copier.CopyCollection(selActivity.Assumptions)
                    End Using
                    Dim ParentPurpose As Purpose = CurrentLogFrame.Purposes(CurrentLogFrame.Purposes.Count - 1)

                    objActivities.Remove(selActivity)
                    ParentPurpose.Outputs.Add(NewOutput)

                    UndoRedo.ItemMovedUp(selActivity, NewOutput, objActivities, intOldIndex, ParentPurpose.Outputs)

                    Return NewOutput
                End If
            Case GetType(Purpose)
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToPreviousParent(selActivity)
        End Select

        Return selActivity
    End Function

    Private Function MoveActivity_ToNextParent(ByVal selActivity As Activity) As LogframeObject
        Dim objActivities As Activities = Me.Logframe.GetParentCollection(selActivity)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim intSection As Integer = selGridRow.Section
        Dim NextStruct As Struct = Me.Grid.GetNextStruct(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objActivities.IndexOf(selActivity)

        If NextStruct Is Nothing Then Return selActivity

        Select Case NextStruct.GetType
            Case GetType(Activity)
                If Me.Logframe.IsParentLineage(NextStruct, selActivity) Then
                    intRowIndex += 1
                    CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                    Return MoveActivity_ToNextParent(selActivity)
                Else
                    Dim objNextActivities As Activities
                    If CType(NextStruct, Activity).Activities.Count > 0 Then
                        'if the next activity has sub-activities, insert as first sub-activity
                        objNextActivities = CType(NextStruct, Activity).Activities
                        intIndex = 0
                    Else
                        'insert before the next activity (at the same level)
                        objNextActivities = Me.Logframe.GetParentCollection(NextStruct)
                        intIndex = objNextActivities.IndexOf(NextStruct)
                    End If

                    objActivities.Remove(selActivity)
                    objNextActivities.Insert(intIndex, selActivity)

                    UndoRedo.ItemMovedDown(selActivity, selActivity, objActivities, intOldIndex, objNextActivities)
                End If

            Case GetType(Output)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToNextParent(selActivity)

            Case GetType(Purpose)
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveActivity_ToNextParent(selActivity)
        End Select

        Return selActivity
    End Function

    Private Function MoveIndicator_ToPreviousParent(ByVal selIndicator As Indicator) As LogframeObject
        Dim objIndicators As Indicators = Me.Logframe.GetParentCollection(selIndicator)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim PreviousIndicator As Indicator = Me.Grid.GetPreviousIndicator(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objIndicators.IndexOf(selIndicator)

        If PreviousIndicator IsNot Nothing Then
            If PreviousIndicator.ParentStructGuid = Guid.Empty And PreviousIndicator.ParentIndicatorGuid = Guid.Empty Then
                'section title row
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToPreviousParent(selIndicator)
            ElseIf Me.Logframe.IsParentLineage(selIndicator, PreviousIndicator) Then
                intRowIndex -= 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToPreviousParent(selIndicator)
            Else
                Dim ParentIndicators As Indicators = Me.Logframe.GetParentCollection(PreviousIndicator)
                intIndex = ParentIndicators.IndexOf(PreviousIndicator)
                If intIndex = ParentIndicators.Count - 1 Then intIndex += 1

                objIndicators.Remove(selIndicator)
                ParentIndicators.Insert(intIndex, selIndicator)

                UndoRedo.ItemMovedUp(selIndicator, selIndicator, objIndicators, intOldIndex, ParentIndicators)
            End If
        End If

        Return selIndicator
    End Function

    Private Function MoveIndicator_ToNextParent(ByVal selIndicator As Indicator) As LogframeObject
        Dim objIndicators As Indicators = Me.Logframe.GetParentCollection(selIndicator)
        Dim intRowIndex As Integer = CurrentCell.RowIndex
        Dim selGridRow As LogframeRow = Me.Grid(intRowIndex)
        Dim NextIndicator As Indicator = Me.Grid.GetNextIndicator(intRowIndex)
        Dim intIndex As Integer
        Dim intOldIndex As Integer = objIndicators.IndexOf(selIndicator)

        If NextIndicator IsNot Nothing Then
            If NextIndicator.ParentStructGuid = Guid.Empty And NextIndicator.ParentIndicatorGuid = Guid.Empty Then
                'section title row
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToNextParent(selIndicator)
            ElseIf Me.Logframe.IsParentLineage(NextIndicator, selIndicator) Then
                intRowIndex += 1
                CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
                Return MoveIndicator_ToNextParent(selIndicator)
            Else
                Dim objNextIndicators As Indicators
                If CType(NextIndicator, Indicator).Indicators.Count > 0 Then
                    'if the next Indicator has sub-Indicators, insert as first sub-Indicator
                    objNextIndicators = CType(NextIndicator, Indicator).Indicators
                    intIndex = 0
                Else
                    'insert before the next Indicator (at the same level)
                    objNextIndicators = Me.Logframe.GetParentCollection(NextIndicator)
                    intIndex = objNextIndicators.IndexOf(NextIndicator)
                End If

                objIndicators.Remove(selIndicator)
                objNextIndicators.Insert(intIndex, selIndicator)

                UndoRedo.ItemMovedDown(selIndicator, selIndicator, objIndicators, intOldIndex, objNextIndicators)
            End If
        ElseIf selGridRow.RowType <> LogframeRow.RowTypes.Normal Then
            intRowIndex += 1
            CurrentCell = Me(CurrentCell.ColumnIndex, intRowIndex)
            Return MoveIndicator_ToNextParent(selIndicator)
        ElseIf IsResourceBudgetRow(selGridRow) Then
            Dim NewResource As New Resource(selIndicator.RTF)
            CType(selGridRow.Struct, Activity).Resources.Insert(0, NewResource)
            objIndicators.Remove(selIndicator)

            Return NewResource
        End If
        Return selIndicator
    End Function

    Private Function MoveResource_ToPreviousParent(ByVal selResource As Resource) As LogframeObject
        Dim objResources As Resources = Me.Logframe.GetParentCollection(selResource)
        Dim intRowIndex As Integer = CurrentCell.RowIndex - 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objResources.IndexOf(selResource)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Struct IsNot Nothing Then
                If IsResourceBudgetRow(selGridRow) Then
                    If selGridRow.Resource IsNot selResource Then
                        CType(selGridRow.Struct, Activity).Resources.Add(selResource)
                        objResources.Remove(selResource)

                        UndoRedo.ItemMovedUp(selResource, selResource, objResources, intOldIndex, CType(selGridRow.Struct, Activity).Resources)
                        Exit Do
                    End If
                Else
                    Dim NewIndicator As New Indicator(selResource.RTF)

                    selGridRow.Struct.Indicators.Add(NewIndicator)
                    objResources.Remove(selResource)

                    UndoRedo.ItemMovedUp(selResource, NewIndicator, objResources, intOldIndex, selGridRow.Struct.Indicators)
                    Return NewIndicator
                End If
            End If
            intRowIndex -= 1
        Loop While intRowIndex >= 0

        Return selResource
    End Function

    Private Sub MoveResource_ToNextParent(ByVal selResource As Resource)
        Dim objResources As Resources = Me.Logframe.GetParentCollection(selResource)
        Dim intRowIndex As Integer = CurrentCell.RowIndex + 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objResources.IndexOf(selResource)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Struct IsNot Nothing Then
                If IsResourceBudgetRow(selGridRow) Then
                    If selGridRow.Resource IsNot selResource And selGridRow.Struct IsNot Nothing Then
                        CType(selGridRow.Struct, Activity).Resources.Insert(0, selResource)
                        objResources.Remove(selResource)

                        UndoRedo.ItemMovedDown(selResource, selResource, objResources, intOldIndex, CType(selGridRow.Struct, Activity).Resources)
                        Exit Do
                    End If
                End If
            End If
            intRowIndex += 1
        Loop While intRowIndex <= Grid.Count - 1
    End Sub

    Private Sub MoveVerificationSource_ToPreviousParent(ByVal selVerificationSource As VerificationSource)
        Dim objVerificationSources As VerificationSources = Me.Logframe.GetParentCollection(selVerificationSource)
        Dim intRowIndex As Integer = CurrentCell.RowIndex - 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objVerificationSources.IndexOf(selVerificationSource)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Indicator IsNot Nothing Then
                If selGridRow.VerificationSource IsNot selVerificationSource Then
                    selGridRow.Indicator.VerificationSources.Add(selVerificationSource)
                    objVerificationSources.Remove(selVerificationSource)

                    UndoRedo.ItemMovedUp(selVerificationSource, selVerificationSource, objVerificationSources, intOldIndex, selGridRow.Indicator.VerificationSources)
                    Exit Do
                End If
            End If
            intRowIndex -= 1
        Loop While intRowIndex >= 0
    End Sub

    Private Sub MoveVerificationSource_ToNextParent(ByVal selVerificationSource As VerificationSource)
        Dim objVerificationSources As VerificationSources = Me.Logframe.GetParentCollection(selVerificationSource)
        Dim intRowIndex As Integer = CurrentCell.RowIndex + 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objVerificationSources.IndexOf(selVerificationSource)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Indicator IsNot Nothing Then
                If IsResourceBudgetRow(selGridRow) = False Then
                    If selGridRow.VerificationSource IsNot selVerificationSource And selGridRow.VerificationSource IsNot Nothing Then
                        selGridRow.Indicator.VerificationSources.Insert(0, selVerificationSource)
                        objVerificationSources.Remove(selVerificationSource)

                        UndoRedo.ItemMovedUp(selVerificationSource, selVerificationSource, objVerificationSources, intOldIndex, selGridRow.Indicator.VerificationSources)
                        Exit Do
                    End If
                End If
            End If
            intRowIndex += 1
        Loop While intRowIndex <= Grid.Count - 1
    End Sub

    Private Sub MoveAssumption_ToPreviousParent(ByVal selAssumption As Assumption)
        Dim objAssumptions As Assumptions = Me.Logframe.GetParentCollection(selAssumption)
        Dim intRowIndex As Integer = CurrentCell.RowIndex - 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objAssumptions.IndexOf(selAssumption)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Struct IsNot Nothing Then
                If selGridRow.Assumption IsNot selAssumption Then
                    selGridRow.Struct.Assumptions.Add(selAssumption)
                    objAssumptions.Remove(selAssumption)

                    UndoRedo.ItemMovedUp(selAssumption, selAssumption, objAssumptions, intOldIndex, selGridRow.Struct.Assumptions)
                    Exit Do
                End If
            End If
            intRowIndex -= 1
        Loop While intRowIndex >= 0
    End Sub

    Private Sub MoveAssumption_ToNextParent(ByVal selAssumption As Assumption)
        Dim objAssumptions As Assumptions = Me.Logframe.GetParentCollection(selAssumption)
        Dim intRowIndex As Integer = CurrentCell.RowIndex + 1
        Dim selGridRow As LogframeRow
        Dim intOldIndex As Integer = objAssumptions.IndexOf(selAssumption)

        Do
            selGridRow = Grid(intRowIndex)
            If selGridRow.RowType = GridRow.RowTypes.Normal And selGridRow.Struct IsNot Nothing Then
                If selGridRow.Assumption IsNot selAssumption And selGridRow.Assumption IsNot Nothing Then
                    selGridRow.Struct.Assumptions.Insert(0, selAssumption)
                    objAssumptions.Remove(selAssumption)

                    UndoRedo.ItemMovedDown(selAssumption, selAssumption, objAssumptions, intOldIndex, selGridRow.Struct.Assumptions)
                    Exit Do
                End If
            End If
            intRowIndex += 1
        Loop While intRowIndex <= Grid.Count - 1
    End Sub

    Public Sub ChangeSection(ByVal intDirection As Integer)
        If CurrentCell Is Nothing Then Exit Sub

        Dim selItem As Struct = Nothing
        Dim selGridRow As LogframeRow = Grid(CurrentCell.RowIndex)
        Dim intIndex, intOldIndex, intParentIndex As Integer

        Select Case selGridRow.Struct.GetType
            Case GetType(Goal)
                Dim selGoal As Goal = CType(selGridRow.Struct, Goal)
                intOldIndex = Me.Logframe.Goals.IndexOf(selGoal)

                If intDirection < 0 Then
                    Exit Sub
                Else
                    Dim NewPurpose As New Purpose(selGoal.RTF)

                    Using copier As New ObjectCopy
                        NewPurpose.Indicators = copier.CopyCollection(selGoal.Indicators)
                        NewPurpose.Assumptions = copier.CopyCollection(selGoal.Assumptions)
                    End Using

                    intIndex = Me.Logframe.Goals.IndexOf(selGoal)
                    If intIndex > Me.Logframe.Purposes.Count - 1 Then intIndex = Me.Logframe.Purposes.Count - 1

                    Me.Logframe.Purposes.Insert(intIndex, NewPurpose)
                    Me.Logframe.Goals.Remove(selGoal)

                    UndoRedo.ItemSectionDown(selGoal, NewPurpose, Me.Logframe.Goals, intOldIndex, Me.Logframe.Purposes)
                    selItem = NewPurpose
                End If
            Case GetType(Purpose)
                Dim selPurpose As Purpose = CType(selGridRow.Struct, Purpose)
                intOldIndex = Me.Logframe.Purposes.IndexOf(selPurpose)

                If intDirection < 0 Then
                    Dim NewGoal As New Goal(selPurpose.RTF)

                    Using copier As New ObjectCopy
                        NewGoal.Indicators = copier.CopyCollection(selPurpose.Indicators)
                        NewGoal.Assumptions = copier.CopyCollection(selPurpose.Assumptions)
                    End Using

                    intIndex = Me.Logframe.Purposes.IndexOf(selPurpose)

                    If selPurpose.Outputs.Count = 0 Then
                        Me.Logframe.Purposes.Remove(selPurpose)
                    Else
                        selPurpose.RTF = String.Empty
                        selPurpose.Indicators.Clear()
                        selPurpose.Assumptions.Clear()
                        selPurpose.TargetGroups.Clear()
                    End If

                    If intIndex > Me.Logframe.Goals.Count - 1 Then intIndex = Me.Logframe.Goals.Count - 1
                    Me.Logframe.Goals.Insert(intIndex, NewGoal)

                    UndoRedo.ItemSectionUp(selPurpose, NewGoal, Me.Logframe.Purposes, intOldIndex, Me.Logframe.Goals)
                    selItem = NewGoal
                Else
                    Dim ParentPurpose As Purpose
                    Dim NewOutput As New Output(selPurpose.RTF)

                    Using copier As New ObjectCopy
                        NewOutput.Indicators = copier.CopyCollection(selPurpose.Indicators)
                        NewOutput.Assumptions = copier.CopyCollection(selPurpose.Assumptions)
                    End Using

                    intIndex = Me.Logframe.Purposes.IndexOf(selPurpose)
                    intParentIndex = intIndex - 1
                    If intParentIndex > Me.Logframe.Purposes.Count - 1 Then intParentIndex = Me.Logframe.Purposes.Count - 1
                    If intParentIndex < 0 Then intParentIndex = 0


                    If selPurpose.Outputs.Count = 0 Then
                        Me.Logframe.Purposes.Remove(selPurpose)
                    Else
                        selPurpose.RTF = String.Empty
                        selPurpose.Indicators.Clear()
                        selPurpose.Assumptions.Clear()
                        selPurpose.TargetGroups.Clear()
                    End If

                    ParentPurpose = Me.Logframe.Purposes(intParentIndex)
                    If intIndex > ParentPurpose.Outputs.Count - 1 Then intIndex = ParentPurpose.Outputs.Count - 1
                    ParentPurpose.Outputs.Insert(intIndex, NewOutput)

                    UndoRedo.ItemSectionDown(selPurpose, NewOutput, Me.Logframe.Purposes, intOldIndex, ParentPurpose.Outputs)

                    selItem = NewOutput
                End If
            Case GetType(Output)
                Dim selOutput As Output = CType(selGridRow.Struct, Output)
                Dim ParentPurpose As Purpose = Me.Logframe.GetParent(selOutput)
                intOldIndex = ParentPurpose.Outputs.IndexOf(selOutput)

                If intDirection < 0 Then
                    Dim NewPurpose As New Purpose(selOutput.RTF)

                    Using copier As New ObjectCopy
                        NewPurpose.Indicators = copier.CopyCollection(selOutput.Indicators)
                        NewPurpose.Assumptions = copier.CopyCollection(selOutput.Assumptions)
                    End Using

                    intIndex = ParentPurpose.Outputs.IndexOf(selOutput)

                    If selOutput.Activities.Count = 0 Then
                        ParentPurpose.Outputs.Remove(selOutput)
                    Else
                        selOutput.RTF = String.Empty
                        selOutput.Indicators.Clear()
                        selOutput.Assumptions.Clear()
                        selOutput.KeyMoments.Clear()
                    End If

                    If intIndex > Me.Logframe.Purposes.Count - 1 Then intIndex = Me.Logframe.Purposes.Count - 1
                    Me.Logframe.Purposes.Insert(intIndex, NewPurpose)

                    UndoRedo.ItemSectionUp(selOutput, NewPurpose, ParentPurpose.Outputs, intOldIndex, Me.Logframe.Purposes)
                    selItem = NewPurpose
                Else
                    Dim ParentOutput As Output
                    Dim NewActivity As New Activity(selOutput.RTF)

                    Using copier As New ObjectCopy
                        NewActivity.Indicators = copier.CopyCollection(selOutput.Indicators)
                        NewActivity.Assumptions = copier.CopyCollection(selOutput.Assumptions)
                    End Using

                    intIndex = ParentPurpose.Outputs.IndexOf(selOutput)
                    intParentIndex = intIndex - 1

                    If intParentIndex > ParentPurpose.Outputs.Count - 1 Then intParentIndex = ParentPurpose.Outputs.Count - 1
                    If intParentIndex < 0 Then intParentIndex = 0

                    If selOutput.Activities.Count = 0 Then
                        ParentPurpose.Outputs.Remove(selOutput)
                    Else
                        selOutput.RTF = String.Empty
                        selOutput.Indicators.Clear()
                        selOutput.Assumptions.Clear()
                        selOutput.KeyMoments.Clear()
                    End If

                    ParentOutput = ParentPurpose.Outputs(intParentIndex)
                    If intIndex > ParentOutput.Activities.Count - 1 Then intIndex = ParentOutput.Activities.Count - 1
                    ParentOutput.Activities.Insert(intIndex, NewActivity)

                    UndoRedo.ItemSectionDown(selOutput, NewActivity, ParentPurpose.Outputs, intOldIndex, ParentOutput.Activities)
                    selItem = NewActivity
                End If
            Case GetType(Activity)
                Dim selActivity As Activity = CType(selGridRow.Struct, Activity)
                Dim ParentActivities As Activities = Me.Logframe.GetParentCollection(selActivity)
                intOldIndex = ParentActivities.IndexOf(selActivity)

                If intDirection < 0 Then
                    Dim ParentPurpose As Purpose

                    Dim NewOutput As New Output(selActivity.RTF)

                    Using copier As New ObjectCopy
                        NewOutput.Indicators = copier.CopyCollection(selActivity.Indicators)
                        NewOutput.Assumptions = copier.CopyCollection(selActivity.Assumptions)
                    End Using

                    intIndex = ParentActivities.IndexOf(selActivity)
                    intParentIndex = intIndex - 1

                    If intParentIndex > Me.Logframe.Purposes.Count - 1 Then intParentIndex = Me.Logframe.Purposes.Count - 1
                    If intParentIndex < 0 Then intParentIndex = 0

                    ParentPurpose = Me.Logframe.Purposes(intParentIndex)
                    If intIndex > ParentPurpose.Outputs.Count - 1 Then intIndex = ParentPurpose.Outputs.Count - 1

                    ParentActivities.Remove(selActivity)
                    ParentPurpose.Outputs.Insert(intIndex, NewOutput)

                    UndoRedo.ItemSectionUp(selActivity, NewOutput, ParentActivities, intOldIndex, ParentPurpose.Outputs)
                    selItem = NewOutput
                Else
                    Exit Sub
                End If
        End Select

        Reload()
        SetFocusOnItem(selItem)
        CurrentRow.Selected = True
    End Sub

    Public Sub LevelUp()
        Dim selItem As LogframeObject = TryCast(Me.CurrentItem(False), LogframeObject)

        Select Case selItem.GetType
            Case GetType(Activity)
                Dim selActivity As Activity = TryCast(Me.CurrentItem(False), Activity)
                If selActivity Is Nothing Then Exit Sub

                Dim ParentActivity As Activity = TryCast(Me.Logframe.GetParent(selActivity), Activity)
                Dim intOldIndex As Integer = ParentActivity.Activities.IndexOf(selActivity)
                If ParentActivity Is Nothing Then Exit Sub

                Dim objParentActivities As Activities = Me.Logframe.GetParentCollection(ParentActivity)
                Dim intIndex As Integer = objParentActivities.IndexOf(ParentActivity)

                intIndex += 1
                ParentActivity.Activities.Remove(selActivity)
                objParentActivities.Insert(intIndex, selActivity)

                UndoRedo.ItemLevelUp(selActivity, ParentActivity.Activities, intOldIndex, objParentActivities)

                Reload()
                SetFocusOnItem(selActivity)
            Case GetType(Indicator)
                Dim selIndicator As Indicator = TryCast(Me.CurrentItem(False), Indicator)
                If selIndicator Is Nothing Then Exit Sub

                Dim ParentIndicator As Indicator = TryCast(Me.Logframe.GetParent(selIndicator), Indicator)
                Dim intOldIndex As Integer = ParentIndicator.Indicators.IndexOf(selIndicator)
                If ParentIndicator Is Nothing Then Exit Sub

                Dim objParentIndicators As Indicators = Me.Logframe.GetParentCollection(ParentIndicator)
                Dim intIndex As Integer = objParentIndicators.IndexOf(ParentIndicator)

                intIndex += 1
                ParentIndicator.Indicators.Remove(selIndicator)
                objParentIndicators.Insert(intIndex, selIndicator) 'IntoProcess

                UndoRedo.ItemLevelUp(selIndicator, ParentIndicator.Indicators, intOldIndex, objParentIndicators)

                Reload()
                SetFocusOnItem(selIndicator)
        End Select
    End Sub

    Public Sub LevelDown()
        Dim selItem As LogframeObject = TryCast(Me.CurrentItem(False), LogframeObject)

        Select Case selItem.GetType
            Case GetType(Activity)
                Dim selActivity As Activity = TryCast(Me.CurrentItem(False), Activity)
                If selActivity Is Nothing Then Exit Sub

                Dim objActivities As Activities = Me.Logframe.GetParentCollection(selActivity)
                Dim intIndex As Integer = objActivities.IndexOf(selActivity)
                Dim intOldIndex As Integer = intIndex
                If intIndex = 0 Then Exit Sub

                intIndex -= 1
                objActivities.Remove(selActivity)
                Dim PreviousActivity As Activity = objActivities(intIndex)
                PreviousActivity.Activities.AddToProcess(selActivity)

                UndoRedo.ItemLevelDown(selActivity, objActivities, intOldIndex, PreviousActivity.Activities)

                Reload()
                SetFocusOnItem(selActivity)
            Case GetType(Indicator)
                Dim selIndicator As Indicator = TryCast(Me.CurrentItem(False), Indicator)
                If selIndicator Is Nothing Then Exit Sub

                Dim objIndicators As Indicators = Me.Logframe.GetParentCollection(selIndicator)
                Dim intIndex As Integer = objIndicators.IndexOf(selIndicator)
                Dim intOldIndex As Integer = intIndex
                If intIndex = 0 Then Exit Sub

                intIndex -= 1
                objIndicators.Remove(selIndicator)
                Dim PreviousIndicator As Indicator = objIndicators(intIndex)
                PreviousIndicator.Indicators.Add(selIndicator)

                UndoRedo.ItemLevelDown(selIndicator, objIndicators, intOldIndex, PreviousIndicator.Indicators)

                Reload()
                SetFocusOnItem(selIndicator)
        End Select
    End Sub
#End Region

#Region "Remove items"
    Public Overrides Sub RemoveItems(ByVal ShowWarning As Boolean, Optional ByVal boolCut As Boolean = False)
        Dim strSourceColName As String
        Dim boolShift As Boolean
        Dim selCell As DataGridViewCell
        Dim boolReload As Boolean
        Dim intRowIndex, intColumnIndex As Integer

        If Me.IsCurrentCellInEditMode = False Then
            selCell = Me(SelectionRectangle.FirstColumnIndex, SelectionRectangle.FirstRowIndex)
            intRowIndex = selCell.RowIndex
            intColumnIndex = selCell.ColumnIndex
            If Control.ModifierKeys = Keys.Shift Then boolShift = True

            'select cells to delete
            strSourceColName = Columns(SelectionRectangle.FirstColumnIndex).Name
            For i = SelectionRectangle.FirstRowIndex To SelectionRectangle.LastRowIndex
                SelectedGridRows.Add(Me.Grid(i))
            Next

            If ShowWarning = True Then
                Dim boolCancel As Boolean = RemoveItems_Warning(strSourceColName)
                If boolCancel = True Then Exit Sub
            End If

            Select Case strSourceColName
                Case "StructSort", "StructRTF"
                    RemoveStructs(boolShift, boolCut)

                Case "IndSort", "IndRTF"
                    If IsResourceBudgetRow(SelectedGridRows(0)) = False Then
                        RemoveIndicators(boolCut)
                    Else
                        RemoveResources(boolCut)
                    End If
                    boolReload = True

                Case "VerSort", "VerRTF"
                    If IsResourceBudgetRow(SelectedGridRows(0)) = False Then
                        RemoveVerificationSources(boolCut)
                        boolReload = True
                    Else
                        RemoveResourceBudgets(boolCut)
                    End If


                Case "AsmSort", "AsmRTF"
                    RemoveAssumptions(boolCut)
            End Select
            Me.RowCount = Me.Grid.Count

            SelectedGridRows.Clear()
            If boolReload = True Then ClearSelection()

            CurrentCell = Me(intColumnIndex, intRowIndex)
            Reload()
        Else
            With CurrentEditingControl
                Dim intSelectionStart As Integer = .SelectionStart
                Dim intLength As Integer = .SelectionLength

                If intLength = 0 Then intLength = 1
                .Text = .Text.Remove(intSelectionStart, intLength)

                .Select(intSelectionStart, 0)
            End With
        End If
    End Sub

    Private Function RemoveItems_Warning(ByVal strSourceColName As String) As Boolean
        Dim objStruct As Struct, objOldStruct As Struct = Nothing
        Dim objIndicator As New Indicator, objOldIndicator As Indicator = Nothing
        Dim objVerificationSource As New VerificationSource
        Dim objResource As New Resource, objOldResource As Resource = Nothing
        Dim objBudgetItem As New BudgetItem
        Dim objAssumption As New Assumption
        Dim intNrInd As Integer, intNrVer As Integer, intNrAsm As Integer
        Dim intNrRsc As Integer, intNrBud As Integer
        Dim intNrRows As Integer
        Dim strMsg As String, strTitle As String = String.Empty
        Dim strParentSing As String = String.Empty, strParentPlur As String = String.Empty
        Dim intTotalNmbr As Integer

        For Each selGridRow As LogframeRow In SelectedGridRows
            objStruct = selGridRow.Struct
            Select Case strSourceColName
                Case "StructSort", "StructRTF"
                    If objStruct IsNot objOldStruct And objStruct IsNot Nothing Then
                        Dim selActivity As Activity = TryCast(objStruct, Activity)

                        intNrInd += objStruct.Indicators.Count
                        If selActivity IsNot Nothing Then intNrRsc += selActivity.Resources.Count
                        intNrAsm += objStruct.Assumptions.Count
                        intTotalNmbr += 1
                    End If
                Case "IndSort", "IndRTF"
                    If IsResourceBudgetRow(selGridRow) = False Then
                        objIndicator = selGridRow.Indicator
                        If objIndicator IsNot objOldIndicator And objIndicator IsNot Nothing Then
                            intNrVer += selGridRow.Indicator.VerificationSources.Count
                            intNrRows = intNrVer
                            intTotalNmbr += 1
                        End If
                        objOldIndicator = objIndicator
                    Else
                        objResource = selGridRow.Resource
                        If objResource IsNot objOldResource And objResource IsNot Nothing Then
                            intNrBud += selGridRow.Resource.BudgetItemReferences.Count
                            intNrRows = intNrBud
                            intTotalNmbr += 1
                        End If
                        objOldResource = objResource
                    End If
            End Select
            objOldStruct = objStruct
        Next
        If intNrInd > 0 Or intNrVer > 0 Or intNrRsc > 0 Or intNrBud > 0 Or intNrAsm > 0 Then
            'modify selection rectangle to include children that will be deleted
            Dim strInd As String = String.Empty, strVer As String = String.Empty, strAsm As String = String.Empty
            Dim strRsc As String = String.Empty, strBud As String = String.Empty
            Dim intLastIndex As Integer = Me.Grid.IndexOf(SelectedGridRows(SelectedGridRows.Count - 1))

            Select Case strSourceColName
                Case "StructSort", "StructRTF"
                    Me("AsmRTF", intLastIndex).Selected = True

                    Select Case SelectedGridRows(0).Section
                        Case Logframe.SectionTypes.GoalsSection
                            strParentSing = My.Settings.setStruct1sing.ToLower
                            strParentPlur = My.Settings.setStruct1.ToLower
                        Case Logframe.SectionTypes.PurposesSection
                            strParentSing = My.Settings.setStruct2sing.ToLower
                            strParentPlur = My.Settings.setStruct2.ToLower
                        Case Logframe.SectionTypes.OutputsSection
                            strParentSing = My.Settings.setStruct3sing.ToLower
                            strParentPlur = My.Settings.setStruct3.ToLower
                        Case Logframe.SectionTypes.ActivitiesSection
                            strParentSing = My.Settings.setStruct4sing.ToLower
                            strParentPlur = My.Settings.setStruct4.ToLower
                    End Select

                    strTitle = LANG_HorizontalDependenciesRemovedTitle
                Case "IndSort", "IndRTF"
                    Me("VerRTF", intLastIndex).Selected = True

                    If IsResourceBudgetRow(SelectedGridRows(0)) = False Then
                        strParentSing = Indicator.ItemName
                        strParentPlur = Indicator.ItemNamePlural

                        strTitle = LANG_VerificationSourcesRemoved
                    Else
                        strParentSing = Resource.ItemName
                        strParentPlur = Resource.ItemNamePlural

                        strTitle = LANG_BudgetItemsRemoved
                    End If
            End Select
            Invalidate(SelectionRectangle.Rectangle) 'SelectionRectangle()

            'show warning
            If intNrInd = 1 Then strInd = intNrInd.ToString & " " & Indicator.ItemName.ToLower
            If intNrInd > 1 Then strInd = intNrInd.ToString & " " & Indicator.ItemNamePlural.ToLower
            If intNrVer = 1 Then strVer = intNrVer.ToString & " " & VerificationSource.ItemName.ToLower
            If intNrVer > 1 Then strVer = intNrVer.ToString & " " & VerificationSource.ItemNamePlural.ToLower
            If intNrRsc = 1 Then strRsc = intNrRsc.ToString & " " & Resource.ItemName.ToLower
            If intNrRsc > 1 Then strRsc = intNrRsc.ToString & " " & Resource.ItemNamePlural.ToLower
            If intNrBud = 1 Then strBud = intNrBud.ToString & " " & BudgetItem.ItemName.ToLower
            If intNrBud > 1 Then strBud = intNrBud.ToString & " " & BudgetItem.ItemNamePlural.ToLower
            If intNrAsm = 1 Then strAsm = intNrAsm.ToString & " " & Assumption.ItemName.ToLower
            If intNrAsm > 1 Then strAsm = intNrAsm.ToString & " " & Assumption.ItemNamePlural.ToLower

            Dim strSumUp1 As String = strInd & IIf(String.IsNullOrEmpty(strInd) = False And String.IsNullOrEmpty(strVer) = False, LANG_And, String.Empty) & strVer
            Dim strSumUp2 As String = strRsc & IIf(String.IsNullOrEmpty(strRsc) = False And String.IsNullOrEmpty(strBud) = False, LANG_And, String.Empty) & strBud
            Dim strSumUp3 As String = strAsm
            Dim strSumUp As String = IIf(String.IsNullOrEmpty(strSumUp1) = False, "  - " & strSumUp1 & vbLf, String.Empty) & _
                IIf(String.IsNullOrEmpty(strSumUp2) = False, "  - " & strSumUp2 & vbLf, String.Empty) & _
                IIf(String.IsNullOrEmpty(strSumUp3) = False, "  - " & strSumUp3 & vbLf, String.Empty)

            If intTotalNmbr = 1 Then
                strMsg = String.Format(LANG_HorizontalDependenciesRemoved, {strParentSing, strSumUp, strParentSing})
            Else
                strMsg = String.Format(LANG_HorizontalDependenciesRemovedPlur, {strParentPlur, strSumUp, strParentPlur})
            End If

            Dim wdDeleteChildren As New DialogWarning(strMsg, strTitle)
            wdDeleteChildren.Type = DialogWarning.WarningDialogTypes.wdDeleteChildren
            wdDeleteChildren.ShowDialog()

            If wdDeleteChildren.DialogResult = Windows.Forms.DialogResult.No Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    Private Sub RemoveStructs(ByVal boolShift As Boolean, ByVal boolCut As Boolean)
        Dim objStruct As Struct, objOldStruct As Struct = Nothing
        Dim RemoveStruct As Struct = Nothing
        Dim boolRemoveAll As Boolean, boolSkip As Boolean
        Dim intObjectType As Integer
        Dim strSortNumber As String = String.Empty
        Dim objParentCollection As Object = Nothing

        For Each selGridRow As LogframeRow In SelectedGridRows
            objStruct = selGridRow.Struct
            If objStruct IsNot Nothing Then intObjectType = objStruct.Section + 10

            If objStruct IsNot objOldStruct And objStruct IsNot Nothing Then
                RemoveStruct = Nothing
                If (objStruct.GetType) Is GetType(Purpose) Then
                    Dim selPurpose As Purpose = CType(objStruct, Purpose)
                    If selPurpose.Outputs.Count > 0 Then
                        If boolShift = True Then boolRemoveAll = True Else boolRemoveAll = False
                    Else
                        boolRemoveAll = True
                    End If
                ElseIf (objStruct.GetType) Is GetType(Output) Then
                    Dim selOutput As Output = CType(objStruct, Output)
                    If selOutput.Activities.Count > 0 Then
                        If boolShift = True Then boolRemoveAll = True Else boolRemoveAll = False
                    Else
                        boolRemoveAll = True
                    End If
                Else
                    boolRemoveAll = True
                End If

                strSortNumber = selGridRow.StructSort
                objParentCollection = Me.Logframe.GetParentCollection(objStruct)

                If boolRemoveAll = True Then
                    If boolCut = False Then
                        UndoRedo.ItemRemoved(objStruct, objParentCollection)
                    Else
                        UndoRedo.ItemCut(objStruct, objParentCollection)
                    End If

                    If objStruct.GetType = GetType(Activity) Then RemoveItems_Referers(objStruct.Guid)
                    Me.Grid.RemoveStruct(objStruct)
                Else
                    If boolCut = False Then
                        UndoRedo.ItemRemovedNotVertical(objStruct, objParentCollection)
                    Else
                        UndoRedo.ItemCutNotVertical(objStruct, objParentCollection)
                    End If

                    objStruct.RTF = String.Empty
                    objStruct.Indicators.Clear()
                    objStruct.Assumptions.Clear()
                    If (objStruct.GetType) Is GetType(Activity) Then CType(objStruct, Activity).Resources.Clear()
                    If (objStruct.GetType) Is GetType(Purpose) Then CType(objStruct, Purpose).TargetGroups.Clear()
                    If (objStruct.GetType) Is GetType(Output) Then CType(objStruct, Output).KeyMoments.Clear()

                    RemoveStruct = objStruct
                    boolSkip = True
                End If

            End If
            If selGridRow.Struct IsNot Nothing And selGridRow.Struct Is RemoveStruct Then
                If boolSkip = False Then
                    Me.Grid.Remove(selGridRow)
                End If
                boolSkip = False
            End If

            objOldStruct = objStruct
        Next
    End Sub

    Private Sub RemoveItems_Referers(ByVal objReferenceGuid As Guid)
        Dim ReferersList As ArrayList = CurrentLogFrame.GetReferingMomentsByReferenceGuid(objReferenceGuid)

        If ReferersList.Count > 0 Then
            For i = 0 To ReferersList.Count - 1
                Select Case ReferersList(i).GetType
                    Case GetType(KeyMoment)
                        Dim selKeyMoment As KeyMoment = ReferersList(i)

                        With selKeyMoment
                            .KeyMoment = .ExactDateKeyMoment
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = KeyMoment.DirectionValues.NotDefined
                            .PeriodUnit = DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                    Case GetType(Activity)
                        Dim selActivity As Activity = ReferersList(i)

                        With selActivity.ActivityDetail
                            .StartDate = .StartDateMainActivity
                            .Relative = False
                            .Period = 0
                            .PeriodDirection = ActivityDetail.DirectionValues.NotDefined
                            .PeriodUnit = ActivityDetail.DurationUnits.NotDefined
                            .GuidReferenceMoment = Guid.Empty
                        End With
                End Select
            Next
        End If
    End Sub

    Private Sub RemoveIndicators(ByVal boolCut As Boolean)
        Dim objInd As Indicator, objOldInd As Indicator = Nothing
        Dim ParentIndicators As Indicators = Nothing
        Dim selGridRow As LogframeRow

        For i = 0 To SelectedGridRows.Count - 1
            selGridRow = SelectedGridRows(i)
            objInd = selGridRow.Indicator

            If objInd IsNot objOldInd And objInd IsNot Nothing Then
                ParentIndicators = Me.Logframe.GetParentCollection(objInd)
                If ParentIndicators IsNot Nothing Then
                    If boolCut = False Then
                        UndoRedo.ItemRemoved(objInd, ParentIndicators)
                    Else
                        UndoRedo.ItemCut(objInd, ParentIndicators)
                    End If

                    ParentIndicators.Remove(objInd)
                End If
            End If
            objOldInd = objInd
        Next
    End Sub

    Private Sub RemoveResources(ByVal boolCut As Boolean)
        Dim objRsc As Resource, objOldRsc As Resource = Nothing
        Dim ParentResources As Resources = Nothing
        Dim selGridRow As LogframeRow

        For i = 0 To SelectedGridRows.Count - 1
            selGridRow = SelectedGridRows(i)
            objRsc = selGridRow.Resource

            If objRsc IsNot objOldRsc And objRsc IsNot Nothing Then
                ParentResources = Me.Logframe.GetParentCollection(objRsc)
                If ParentResources IsNot Nothing Then
                    If boolCut = False Then
                        UndoRedo.ItemRemoved(objRsc, ParentResources)
                    Else
                        UndoRedo.ItemCut(objRsc, ParentResources)
                    End If

                    ParentResources.Remove(objRsc)
                End If
            End If
            objOldRsc = objRsc
        Next
    End Sub

    Private Sub RemoveVerificationSources(ByVal boolCut As Boolean)
        Dim objVer As VerificationSource, objOldVer As VerificationSource = Nothing
        Dim ParentVerificationSources As VerificationSources = Nothing
        Dim selGridRow As LogframeRow

        For i = 0 To SelectedGridRows.Count - 1
            selGridRow = SelectedGridRows(i)
            objVer = selGridRow.VerificationSource

            If objVer IsNot objOldVer And objVer IsNot Nothing Then
                ParentVerificationSources = Me.Logframe.GetParentCollection(objVer)
                If ParentVerificationSources IsNot Nothing Then
                    If boolCut = False Then
                        UndoRedo.ItemRemoved(objVer, ParentVerificationSources)
                    Else
                        UndoRedo.ItemCut(objVer, ParentVerificationSources)
                    End If
                    ParentVerificationSources.Remove(objVer)
                End If
            End If
            objOldVer = objVer
        Next
    End Sub

    Private Sub RemoveResourceBudgets(ByVal boolCut As Boolean)
        Dim objRsc As Resource, objOldRsc As Resource = Nothing
        Dim ParentResources As Resources = Nothing
        Dim selGridRow As LogframeRow

        For i = 0 To SelectedGridRows.Count - 1
            selGridRow = SelectedGridRows(i)
            objRsc = selGridRow.Resource

            If objRsc IsNot objOldRsc And objRsc IsNot Nothing Then
                If objRsc.TotalCostAmount > 0 Then
                    objRsc.BudgetItemReferences.Clear()
                    objRsc.SetTotalCostAmount()
                End If

            End If
            objOldRsc = objRsc
        Next
    End Sub

    Private Sub RemoveAssumptions(ByVal boolCut As Boolean)
        Dim objAsm As Assumption, objOldAsm As Assumption = Nothing
        Dim ParentAssumptions As Assumptions = Nothing
        Dim selGridRow As LogframeRow

        For i = 0 To SelectedGridRows.Count - 1
            selGridRow = SelectedGridRows(i)
            objAsm = selGridRow.Assumption

            If objAsm IsNot objOldAsm And objAsm IsNot Nothing Then
                ParentAssumptions = Me.Logframe.GetParentCollection(objAsm)
                If ParentAssumptions IsNot Nothing Then
                    If boolCut = False Then
                        UndoRedo.ItemRemoved(objAsm, ParentAssumptions)
                    Else
                        UndoRedo.ItemCut(objAsm, ParentAssumptions)
                    End If
                    ParentAssumptions.Remove(objAsm)
                End If
            End If
            objOldAsm = objAsm
        Next
    End Sub

#End Region 'Remove items

#Region "Copy and paste cells"
    Public Overrides Sub CutItems(ByVal ShowWarning As Boolean)
        CopyItems()

        RemoveItems(ShowWarning, True)
    End Sub

    Public Overrides Sub CopyItems()
        With SelectionRectangle
            Dim strFirstColumnName As String = Me.Columns(.FirstColumnIndex).Name
            Dim strLastColumnName As String = Me.Columns(.LastColumnIndex).Name
            Dim intChildren As Integer
            Dim CopyGroup As Date = Now()

            Select Case strFirstColumnName
                Case "StructSort", "StructRTF"
                    If strLastColumnName.Contains("Ind") Then
                        intChildren = 1
                    ElseIf strLastColumnName.Contains("Ver") Then
                        intChildren = 3 '+1 for ind and +2 for ver
                    ElseIf strLastColumnName.Contains("Asm") Then
                        intChildren = 4
                        If ShowIndicatorColumn = True Then intChildren += 1
                        If ShowVerificationSourceColumn = True Then intChildren += 2
                    End If
                    CopyItems_Copy(strFirstColumnName, CopyGroup, intChildren)
                Case "IndSort", "IndRTF"
                    If strLastColumnName.Contains("Ind") Then
                        CopyItems_Copy(strFirstColumnName, CopyGroup, intChildren)
                    ElseIf strLastColumnName.Contains("Ver") Then
                        intChildren = 1
                        CopyItems_Copy(strFirstColumnName, CopyGroup, intChildren)
                    ElseIf strLastColumnName.Contains("Asm") Then
                        intChildren = 1
                        CopyItems_Copy(strFirstColumnName, CopyGroup, intChildren)
                        CopyItems_Copy(strLastColumnName, CopyGroup, 0)
                    End If
                Case "VerSort", "VerRTF"
                    If strLastColumnName.Contains("Ver") Then
                        CopyItems_Copy(strFirstColumnName, CopyGroup, intChildren)
                    ElseIf strLastColumnName.Contains("Asm") Then
                        CopyItems_Copy(strFirstColumnName, CopyGroup, 0)
                        CopyItems_Copy(strLastColumnName, CopyGroup, 0)
                    End If
                Case "AsmSort", "AsmRTF"
                    CopyItems_Copy(strFirstColumnName, CopyGroup, 0)
            End Select
        End With

    End Sub

    Private Sub CopyItems_Copy(ByVal strColumnName As String, ByVal CopyGroup As Date, ByVal intChildren As Integer)
        Dim selRow As LogframeRow
        Dim strSort As String

        With SelectionRectangle
            Select Case strColumnName
                Case "StructSort", "StructRTF"
                    For i = .FirstRowIndex To .LastRowIndex
                        selRow = Me.Grid(i)
                        If String.IsNullOrEmpty(selRow.StructRtf) = False Then
                            If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Struct Then
                                strSort = selRow.Struct.GetItemName(selRow.Section) & " " & selRow.StructSort
                                Dim NewItem As New ClipboardItem(CopyGroup, selRow.Struct, strSort, intChildren)
                                ItemClipboard.Insert(0, NewItem)
                            End If
                        End If
                    Next
                Case "IndSort", "IndRTF"
                    For i = .FirstRowIndex To .LastRowIndex
                        selRow = Me.Grid(i)
                        If IsResourceBudgetRow(selRow) = False Then
                            If String.IsNullOrEmpty(selRow.IndicatorRtf) = False Then
                                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Indicator Then
                                    strSort = Indicator.ItemName & " " & selRow.IndicatorSort
                                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.Indicator, strSort, intChildren)
                                    ItemClipboard.Insert(0, NewItem)
                                End If
                            End If
                        Else
                            If String.IsNullOrEmpty(selRow.ResourceRtf) = False Then
                                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Resource Then
                                    strSort = Resource.ItemName & " " & selRow.ResourceSort
                                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.Resource, strSort, intChildren)
                                    ItemClipboard.Insert(0, NewItem)
                                End If
                            End If
                        End If
                    Next
                Case "VerSort", "VerRTF"
                    For i = .FirstRowIndex To .LastRowIndex
                        selRow = Me.Grid(i)
                        If IsResourceBudgetRow(selRow) = False Then
                            If String.IsNullOrEmpty(selRow.VerificationSourceRtf) = False Then
                                If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.VerificationSource Then
                                    strSort = VerificationSource.ItemName & " " & selRow.VerificationSourceSort
                                    Dim NewItem As New ClipboardItem(CopyGroup, selRow.VerificationSource, strSort)
                                    ItemClipboard.Insert(0, NewItem)
                                End If
                            End If
                        Else

                        End If
                    Next
                Case "AsmSort", "AsmRTF"
                    For i = .FirstRowIndex To .LastRowIndex
                        selRow = Me.Grid(i)
                        If String.IsNullOrEmpty(selRow.AssumptionRtf) = False Then
                            If ItemClipboard.Count = 0 OrElse ItemClipboard(0).Item IsNot selRow.Assumption Then
                                strSort = Assumption.ItemName & " " & selRow.AssumptionSort
                                Dim NewItem As New ClipboardItem(CopyGroup, selRow.Assumption, strSort)
                                ItemClipboard.Insert(0, NewItem)
                            End If
                        End If
                    Next
            End Select
        End With
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems, ByVal intPasteOption As Integer, Optional ByVal PasteCell As DataGridViewCell = Nothing)
        If PasteCell Is Nothing Then PasteCell = CurrentCell
        If PasteCell Is Nothing Then Exit Sub

        Dim intColumnIndex As Integer = PasteCell.ColumnIndex
        Dim intRowIndex As Integer = PasteCell.RowIndex

        PasteRow = Grid(intRowIndex)
        Dim strPasteColumnName As String = Me.Columns(PasteCell.ColumnIndex).Name
        Dim selItem As ClipboardItem
        If PasteRow.RowType <> GridRow.RowTypes.Normal Then PasteRow = Me.Grid(PasteCell.RowIndex - 1)

        For i = 0 To PasteItems.Count - 1 'To 0 Step -1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(Goal)
                    PasteItems_Goal(selItem, strPasteColumnName, intPasteOption)
                Case GetType(Purpose)
                    PasteItems_Purpose(selItem, strPasteColumnName, intPasteOption)
                Case GetType(Output)
                    PasteItems_Output(selItem, strPasteColumnName, intPasteOption)
                Case GetType(Activity)
                    PasteItems_Activity(selItem, strPasteColumnName, intPasteOption)
                Case GetType(Indicator)
                    PasteItems_Indicator(selItem, strPasteColumnName, intPasteOption)
                Case GetType(Resource)
                    PasteItems_Resource(selItem, strPasteColumnName, intPasteOption)
                Case GetType(VerificationSource)
                    PasteItems_VerificationSource(selItem, strPasteColumnName)
                Case GetType(Assumption)
                    PasteItems_Assumption(selItem, strPasteColumnName)
            End Select
        Next

        Me.Reload()
        Me.CurrentCell = Me(strPasteColumnName, intRowIndex)
    End Sub

    Private Function PasteItems_GetIndexOfGoal() As Integer
        Dim intIndex As Integer = -1

        If PasteRow.RowType = GridRow.RowTypes.Section Then
            intIndex = CurrentLogFrame.Goals.Count
        Else
            If PasteRow.Struct IsNot Nothing Then
                intIndex = CurrentLogFrame.Goals.IndexOf(CType(PasteRow.Struct, Goal))
            End If
        End If
        If intIndex = -1 Then intIndex = CurrentLogFrame.Goals.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfPurpose() As Integer
        Dim intIndex As Integer = -1

        If PasteRow.RowType = GridRow.RowTypes.Section Then
            intIndex = CurrentLogFrame.Purposes.Count
        Else
            If PasteRow.Struct IsNot Nothing Then
                intIndex = CurrentLogFrame.Purposes.IndexOf(CType(PasteRow.Struct, Purpose))
            End If
        End If
        If intIndex = -1 Then intIndex = CurrentLogFrame.Purposes.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfOutput(ByVal ParentPurpose As Purpose) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.Struct IsNot Nothing Then
            intIndex = ParentPurpose.Outputs.IndexOf(CType(PasteRow.Struct, Output))
        End If
        If intIndex = -1 Then intIndex = ParentPurpose.Outputs.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfActivity(ByVal ParentActivities As Activities) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.Struct IsNot Nothing Then
            intIndex = ParentActivities.IndexOf(CType(PasteRow.Struct, Activity))
        End If
        If intIndex = -1 Then intIndex = ParentActivities.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfIndicator(ByVal ParentIndicators As Indicators) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.Indicator IsNot Nothing Then
            intIndex = ParentIndicators.IndexOf(PasteRow.Indicator)
        End If
        If intIndex = -1 Then intIndex = ParentIndicators.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfResource(ByVal ParentResources As Resources) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.Resource IsNot Nothing Then
            intIndex = ParentResources.IndexOf(PasteRow.Resource)
        End If
        If intIndex = -1 Then intIndex = ParentResources.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfVerificationSource(ByVal ParentVerificationSources As VerificationSources) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.VerificationSource IsNot Nothing Then
            intIndex = ParentVerificationSources.IndexOf(PasteRow.VerificationSource)
        End If
        If intIndex = -1 Then intIndex = ParentVerificationSources.Count

        Return intIndex
    End Function

    Private Function PasteItems_GetIndexOfAssumption(ByVal ParentAssumptions As Assumptions) As Integer
        Dim intIndex As Integer = -1

        If PasteRow.Assumption IsNot Nothing Then
            intIndex = ParentAssumptions.IndexOf(PasteRow.Assumption)
        End If
        If intIndex = -1 Then intIndex = ParentAssumptions.Count

        Return intIndex
    End Function
#End Region

#Region "Paste structs"
    Private Sub PasteItems_Goal(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)
        Dim selGoal As Goal = CType(selItem.Item, Goal)
        Dim strSortNumber As String = PasteRow.StructSort

        Select Case strPasteColumnName
            Case "StructSort", "StructRTF"
                If PasteRow.Section = Logframe.SectionTypes.GoalsSection Then
                    Dim NewGoal As New Goal
                    Dim intIndex As Integer = PasteItems_GetIndexOfGoal()

                    Using copier As New ObjectCopy
                        NewGoal = copier.CopyObject(selGoal)
                    End Using

                    PasteItems_Struct_HorChildren(NewGoal, selItem)

                    CurrentLogFrame.Goals.Insert(intIndex, NewGoal)
                    PasteRow.Struct = NewGoal

                    UndoRedo.ItemPasted(NewGoal, CurrentLogFrame.Goals)
                Else
                    PasteItems_AsStruct(selItem, strPasteColumnName)
                End If
            Case Else
                PasteItems_FromStruct(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Sub PasteItems_Purpose(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)
        Dim selPurpose As Purpose = CType(selItem.Item, Purpose)
        Dim strSortNumber As String = PasteRow.StructSort

        Select Case strPasteColumnName
            Case "StructSort", "StructRTF"
                If PasteRow.Section = Logframe.SectionTypes.PurposesSection Then
                    Dim NewPurpose As New Purpose
                    Dim intIndex As Integer = PasteItems_GetIndexOfPurpose()

                    Using copier As New ObjectCopy
                        NewPurpose = copier.CopyObject(selPurpose)
                    End Using

                    PasteItems_Struct_HorChildren(NewPurpose, selItem)
                    PasteItems_Struct_Dependencies(NewPurpose, intPasteOption)

                    CurrentLogFrame.Purposes.Insert(intIndex, NewPurpose)
                    PasteRow.Struct = NewPurpose

                    UndoRedo.ItemPasted(NewPurpose, CurrentLogFrame.Purposes)
                Else
                    PasteItems_AsStruct(selItem, strPasteColumnName)
                End If
            Case Else
                PasteItems_FromStruct(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Sub PasteItems_Output(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)
        Dim selOutput As Output = CType(selItem.Item, Output)
        Dim strSortNumber As String = PasteRow.StructSort

        Select Case strPasteColumnName
            Case "StructSort", "StructRTF"
                If PasteRow.Section = Logframe.SectionTypes.OutputsSection Then
                    Dim NewOutput As New Output
                    Dim ParentPurpose As Purpose = PasteItems_Output_GetParentPurpose()
                    Dim intIndex As Integer = PasteItems_GetIndexOfOutput(ParentPurpose)

                    Using copier As New ObjectCopy
                        NewOutput = copier.CopyObject(selOutput)
                    End Using

                    PasteItems_Struct_HorChildren(NewOutput, selItem)
                    PasteItems_Struct_Dependencies(NewOutput, intPasteOption)

                    ParentPurpose.Outputs.Insert(intIndex, NewOutput)
                    PasteRow.Struct = NewOutput

                    UndoRedo.ItemPasted(NewOutput, ParentPurpose.Outputs)
                Else
                    PasteItems_AsStruct(selItem, strPasteColumnName)
                End If
            Case Else
                PasteItems_FromStruct(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Function PasteItems_Output_GetParentPurpose() As Purpose
        Dim ParentPurpose As Purpose = Nothing
        If PasteRow.Struct IsNot Nothing Then
            ParentPurpose = Me.Logframe.GetParent(CType(PasteRow.Struct, Output))
        Else
            Dim intRowIndex As Integer = Me.Grid.IndexOf(PasteRow)
            Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(intRowIndex)

            If PreviousStruct Is Nothing Then
                ParentPurpose = Me.Logframe.Purposes(0)
            Else
                Select Case PreviousStruct.GetType
                    Case GetType(Output)
                        ParentPurpose = Me.Logframe.GetParent(CType(PreviousStruct, Output))
                    Case GetType(Purpose) 'Repeated purpose
                        ParentPurpose = PreviousStruct
                End Select
            End If
        End If
        If ParentPurpose Is Nothing Then ParentPurpose = CurrentLogFrame.Purposes(CurrentLogFrame.Purposes.Count - 1)

        Return ParentPurpose
    End Function

    Private Sub PasteItems_Activity(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)
        Dim selActivity As Activity = CType(selItem.Item, Activity)
        Dim strSortNumber As String = PasteRow.StructSort

        Select Case strPasteColumnName
            Case "StructSort", "StructRTF"
                If PasteRow.Section = Logframe.SectionTypes.ActivitiesSection Then
                    Dim NewActivity As New Activity
                    Dim ParentActivities As Activities = PasteItems_Activity_GetParentActivities()
                    Dim intIndex As Integer = PasteItems_GetIndexOfActivity(ParentActivities)

                    Using copier As New ObjectCopy
                        NewActivity = copier.CopyObject(selActivity)
                    End Using

                    PasteItems_Struct_HorChildren(NewActivity, selItem)

                    ParentActivities.Insert(intIndex, NewActivity)
                    PasteRow.Struct = NewActivity

                    UndoRedo.ItemPasted(NewActivity, ParentActivities)
                Else
                    PasteItems_AsStruct(selItem, strPasteColumnName)
                End If
            Case Else
                PasteItems_FromStruct(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Function PasteItems_Activity_GetParentActivities() As Activities
        Dim ParentActivities As Activities = Nothing
        Dim ParentOutput As Output = Nothing
        Dim ParentPurpose As Purpose = Nothing

        If PasteRow.Struct IsNot Nothing Then
            ParentActivities = Me.Logframe.GetParentCollection(CType(PasteRow.Struct, Activity))
        Else
            Dim PreviousStruct As Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

            If PreviousStruct IsNot Nothing Then
                If PreviousStruct.GetType Is GetType(Purpose) Then
                    ParentPurpose = CType(PreviousStruct, Purpose)
                    If ParentPurpose.Outputs.Count = 0 Then ParentPurpose.Outputs.Add(New Output)
                    ParentOutput = CType(PreviousStruct, Purpose).Outputs(0)
                    ParentActivities = ParentOutput.Activities
                ElseIf PreviousStruct.GetType Is GetType(Output) Then
                    ParentOutput = PreviousStruct
                    ParentActivities = ParentOutput.Activities
                Else
                    ParentActivities = Me.Logframe.GetParentCollection(CType(PreviousStruct, Activity))
                End If
            End If
        End If

        If ParentActivities Is Nothing Then
            If ParentPurpose Is Nothing Then ParentPurpose = CurrentLogFrame.Purposes(CurrentLogFrame.Purposes.Count - 1)
            ParentOutput = ParentPurpose.Outputs(ParentPurpose.Outputs.Count - 1)
            ParentActivities = ParentOutput.Activities
        End If

        Return ParentActivities
    End Function

    Private Sub PasteItems_AsStruct(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Dim strSortNumber As String = PasteRow.StructSort
        Dim intIndex As Integer

        Select Case PasteRow.Section
            Case Logframe.SectionTypes.GoalsSection
                Dim NewGoal As New Goal(CType(selItem.Item, LogframeObject).RTF)

                intIndex = PasteItems_GetIndexOfGoal()

                Select Case selItem.Item.GetType
                    Case GetType(Purpose), GetType(Output), GetType(Activity)
                        PasteItems_AsStruct_HorChildren(NewGoal, selItem)
                End Select

                CurrentLogFrame.Goals.Insert(intIndex, NewGoal)
                PasteRow.Struct = NewGoal

                UndoRedo.ItemPasted(NewGoal, CurrentLogFrame.Goals)
            Case Logframe.SectionTypes.PurposesSection
                Dim NewPurpose As New Purpose(CType(selItem.Item, LogframeObject).RTF)

                intIndex = PasteItems_GetIndexOfPurpose()

                Select Case selItem.Item.GetType
                    Case GetType(Goal), GetType(Output), GetType(Activity)
                        PasteItems_AsStruct_HorChildren(NewPurpose, selItem)
                End Select

                CurrentLogFrame.Purposes.Insert(intIndex, NewPurpose)
                PasteRow.Struct = NewPurpose

                UndoRedo.ItemPasted(NewPurpose, CurrentLogFrame.Purposes)
            Case Logframe.SectionTypes.OutputsSection
                Dim NewOutput As New Output(CType(selItem.Item, LogframeObject).RTF)
                Dim ParentPurpose As Purpose = PasteItems_Output_GetParentPurpose()

                intIndex = PasteItems_GetIndexOfOutput(ParentPurpose)

                Select Case selItem.Item.GetType
                    Case GetType(Goal), GetType(Purpose), GetType(Activity)
                        PasteItems_AsStruct_HorChildren(NewOutput, selItem)
                End Select

                ParentPurpose.Outputs.Insert(intIndex, NewOutput)
                PasteRow.Struct = NewOutput

                UndoRedo.ItemPasted(NewOutput, ParentPurpose.Outputs)
            Case Logframe.SectionTypes.ActivitiesSection
                Dim NewActivity As New Activity(CType(selItem.Item, LogframeObject).RTF)
                Dim ParentActivities As Activities = PasteItems_Activity_GetParentActivities()

                intIndex = PasteItems_GetIndexOfActivity(ParentActivities)

                Select Case selItem.Item.GetType
                    Case GetType(Goal), GetType(Output), GetType(Activity)
                        PasteItems_AsStruct_HorChildren(NewActivity, selItem)
                End Select

                ParentActivities.Insert(intIndex, NewActivity)
                PasteRow.Struct = NewActivity

                UndoRedo.ItemPasted(NewActivity, ParentActivities)
        End Select
    End Sub

    Private Sub PasteItems_FromStruct(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Select Case strPasteColumnName
            Case "IndSort", "IndRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsIndicator(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsVerificationSource(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "AsmSort", "AsmRTF"
                PasteItems_AsAssumption(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Sub PasteItems_Struct_HorChildren(ByRef selStruct As Struct, ByVal selItem As ClipboardItem)
        Dim intColumns As Integer = selItem.Children
        If intColumns >= 4 Then
            intColumns -= 4
        Else
            selStruct.Assumptions.Clear()
        End If
        If intColumns = 0 Then
            selStruct.Indicators.Clear()
        ElseIf intColumns = 1 Then
            For Each selIndicator As Indicator In selStruct.Indicators
                selIndicator.VerificationSources.Clear()
            Next
        End If
    End Sub

    Private Sub PasteItems_AsStruct_HorChildren(ByRef selStruct As Struct, ByVal selItem As ClipboardItem)
        Dim intColumns As Integer = selItem.Children

        Using copier As New ObjectCopy
            If intColumns >= 4 Then
                selStruct.Assumptions = copier.CopyCollection(CType(selItem.Item, Struct).Assumptions)
                intColumns -= 4
            End If
            If intColumns > 0 Then selStruct.Indicators = copier.CopyCollection(CType(selItem.Item, Struct).Indicators)
            If intColumns < 3 Then
                For Each selIndicator As Indicator In selStruct.Indicators
                    selIndicator.VerificationSources.Clear()
                Next
            End If
        End Using
    End Sub

    Private Sub PasteItems_Struct_Dependencies(ByRef selStruct As Struct, ByVal intPasteOption As Integer)
        Select Case selStruct.GetType
            Case GetType(Purpose)
                Select Case intPasteOption
                    Case ClipboardItems.PasteOptions.PasteNoVert
                        CType(selStruct, Purpose).Outputs.Clear()
                    Case ClipboardItems.PasteOptions.PasteNoDetails
                        CType(selStruct, Purpose).TargetGroups.Clear()
                    Case ClipboardItems.PasteOptions.PasteNoDependencies
                        CType(selStruct, Purpose).Outputs.Clear()
                        CType(selStruct, Purpose).TargetGroups.Clear()
                End Select
            Case GetType(Output)
                Select Case intPasteOption
                    Case ClipboardItems.PasteOptions.PasteNoVert
                        CType(selStruct, Output).Activities.Clear()
                    Case ClipboardItems.PasteOptions.PasteNoDetails
                        CType(selStruct, Output).KeyMoments.Clear()
                    Case ClipboardItems.PasteOptions.PasteNoDependencies
                        CType(selStruct, Output).Activities.Clear()
                        CType(selStruct, Output).KeyMoments.Clear()
                End Select
            Case GetType(Activity)
                Select Case intPasteOption
                    Case ClipboardItems.PasteOptions.PasteNoDetails, ClipboardItems.PasteOptions.PasteNoDependencies
                        CType(selStruct, Activity).ActivityDetail = New ActivityDetail
                End Select
        End Select
    End Sub
#End Region

#Region "Paste indicators"
    Private Sub PasteItems_Indicator(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)
        If PasteRow.Struct Is Nothing And strPasteColumnName.Contains("Struct") = False Then PasteRow.Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

        Dim selIndicator As Indicator = CType(selItem.Item, Indicator)
        Dim strSortNumber As String = PasteRow.IndicatorSort

        Select Case strPasteColumnName
            Case "IndSort", "IndRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    Dim NewIndicator As New Indicator
                    Dim ParentIndicators As Indicators = PasteItems_Indicator_GetParentIndicators()
                    Dim intIndex As Integer = PasteItems_GetIndexOfIndicator(ParentIndicators)

                    Using copier As New ObjectCopy
                        NewIndicator = copier.CopyObject(selIndicator)
                    End Using

                    PasteItems_Indicator_HorChildren(NewIndicator, selItem)

                    ParentIndicators.Insert(intIndex, NewIndicator)
                    PasteRow.Indicator = NewIndicator

                    UndoRedo.ItemPasted(NewIndicator, ParentIndicators)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "StructSort", "StructRTF"
                PasteItems_AsStruct(selItem, strPasteColumnName)
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsVerificationSource(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "AsmSort", "AsmRTF"
                PasteItems_AsAssumption(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Function PasteItems_Indicator_GetParentIndicators() As Indicators
        Dim ParentIndicators As Indicators = Nothing
        Dim intRowIndex As Integer = Grid.IndexOf(PasteRow)

        If PasteRow.Indicator IsNot Nothing Then
            ParentIndicators = Me.Logframe.GetParentCollection(PasteRow.Indicator)
        Else
            Dim selGridRow As LogframeRow

            For i = intRowIndex To 0 Step -1
                selGridRow = Me.Grid(i)
                If selGridRow.Indicator IsNot Nothing Then
                    ParentIndicators = Me.Logframe.GetParentCollection(selGridRow.Indicator)
                    Exit For
                ElseIf selGridRow.Struct IsNot Nothing Then
                    ParentIndicators = selGridRow.Struct.Indicators
                    Exit For
                End If
            Next
        End If

        If ParentIndicators Is Nothing Then
            Select Case PasteRow.Section
                Case LogframeObject.SectionTypes.Goal
                    Dim NewGoal As New Goal
                    Logframe.Goals.Add(NewGoal)
                    ParentIndicators = NewGoal.Indicators
                Case LogframeObject.SectionTypes.Purpose
                    Dim NewPurpose As New Purpose
                    Logframe.Purposes.Add(NewPurpose)
                    ParentIndicators = NewPurpose.Indicators
                Case LogframeObject.SectionTypes.Output
                    Dim NewOutput As New Output

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)
                    ParentPurpose.Outputs.Add(NewOutput)
                    ParentIndicators = NewOutput.Indicators
                Case LogframeObject.SectionTypes.Activity
                    Dim NewActivity As New Activity

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)

                    If ParentPurpose.Outputs.Count = 0 Then
                        Dim NewOutput As New Output
                        ParentPurpose.Outputs.Add(NewOutput)
                    End If

                    Dim ParentOutput As Output = ParentPurpose.Outputs(ParentPurpose.Outputs.Count - 1)
                    ParentOutput.Activities.Add(NewActivity)
                    ParentIndicators = NewActivity.Indicators
            End Select
        End If

        Return ParentIndicators
    End Function

    Private Sub PasteItems_Indicator_HorChildren(ByRef selIndicator As Indicator, ByVal selItem As ClipboardItem)
        If selItem.Children < 1 Then selIndicator.VerificationSources.Clear()
    End Sub

    Private Sub PasteItems_Indicator_Dependencies(ByRef selIndicator As Indicator, ByVal intPasteOption As Integer)
        Select Case intPasteOption
            Case ClipboardItems.PasteOptions.PasteNoDetails, ClipboardItems.PasteOptions.PasteNoDependencies
                selIndicator.Statements.Clear()
                selIndicator.ResponseClasses.Clear()
        End Select
    End Sub

    Private Sub PasteItems_AsIndicator(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Dim strSortNumber As String = PasteRow.IndicatorSort
        Dim NewIndicator As New Indicator(CType(selItem.Item, LogframeObject).RTF)
        Dim ParentIndicators As Indicators = PasteItems_Indicator_GetParentIndicators()

        If ParentIndicators IsNot Nothing Then
            Dim intIndex As Integer = PasteItems_GetIndexOfIndicator(ParentIndicators)

            ParentIndicators.Insert(intIndex, NewIndicator)
            PasteRow.Indicator = NewIndicator

            UndoRedo.ItemPasted(NewIndicator, ParentIndicators)
        End If
    End Sub
#End Region

#Region "Paste resources"
    Private Sub PasteItems_Resource(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String, ByVal intPasteOption As Integer)

        If PasteRow.Struct Is Nothing And strPasteColumnName.Contains("Struct") = False Then PasteRow.Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

        Dim selResource As Resource = CType(selItem.Item, Resource)
        Dim strSortNumber As String = PasteRow.ResourceSort

        Select Case strPasteColumnName
            Case "IndSort", "IndRTF"
                If IsResourceBudgetRow(PasteRow) Then
                    Dim NewResource As New Resource
                    Dim ParentResources As Resources = PasteItems_Resource_GetParentResources()
                    Dim intIndex As Integer = PasteItems_GetIndexOfResource(ParentResources)

                    Using copier As New ObjectCopy
                        NewResource = copier.CopyObject(selResource)
                    End Using

                    PasteItems_Resource_Dependencies(NewResource, intPasteOption)

                    ParentResources.Insert(intIndex, NewResource)
                    PasteRow.Resource = NewResource

                    UndoRedo.ItemPasted(NewResource, ParentResources)
                Else
                    PasteItems_AsIndicator(selItem, strPasteColumnName)
                End If
            Case "StructSort", "StructRTF"
                PasteItems_AsStruct(selItem, strPasteColumnName)
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsVerificationSource(selItem, strPasteColumnName)
                Else
                    PasteItems_Resource(selItem, "IndRTF", intPasteOption)
                End If
            Case "AsmSort", "AsmRTF"
                PasteItems_AsAssumption(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Function PasteItems_Resource_GetParentResources() As Resources
        Dim ParentResources As Resources = Nothing
        Dim intRowIndex As Integer = Grid.IndexOf(PasteRow)

        If PasteRow.Resource IsNot Nothing Then
            ParentResources = Me.Logframe.GetParentCollection(PasteRow.Resource)
        Else
            Dim selGridRow As LogframeRow

            For i = intRowIndex To 0 Step -1
                selGridRow = Me.Grid(i)
                If selGridRow.Resource IsNot Nothing Then
                    ParentResources = Me.Logframe.GetParentCollection(selGridRow.Resource)
                    Exit For
                ElseIf selGridRow.Struct IsNot Nothing Then
                    ParentResources = CType(selGridRow.Struct, Activity).Resources
                    Exit For
                End If
            Next
        End If

        If ParentResources Is Nothing Then
            Dim NewActivity As New Activity

            If Logframe.Purposes.Count = 0 Then
                Dim NewPurpose As New Purpose
                Logframe.Purposes.Add(NewPurpose)
            End If

            Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)

            If ParentPurpose.Outputs.Count = 0 Then
                Dim NewOutput As New Output
                ParentPurpose.Outputs.Add(NewOutput)
            End If

            Dim ParentOutput As Output = ParentPurpose.Outputs(ParentPurpose.Outputs.Count - 1)
            ParentOutput.Activities.Add(NewActivity)
            ParentResources = NewActivity.Resources
        End If

        Return ParentResources
    End Function

    Private Sub PasteItems_Resource_Dependencies(ByRef selResource As Resource, ByVal intPasteOption As Integer)
        Select Case intPasteOption
            Case ClipboardItems.PasteOptions.PasteNoDetails, ClipboardItems.PasteOptions.PasteNoDependencies
                selResource.BudgetItemReferences.Clear()
                selResource.TotalCostAmount = 0
        End Select
    End Sub

    Private Sub PasteItems_AsResource(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Dim strSortNumber As String = PasteRow.ResourceSort
        Dim NewResource As New Resource(CType(selItem.Item, LogframeObject).RTF)
        Dim ParentResources As Resources = PasteItems_Resource_GetParentResources()

        If ParentResources IsNot Nothing Then
            Dim intIndex As Integer = PasteItems_GetIndexOfResource(ParentResources)

            ParentResources.Insert(intIndex, NewResource)
            PasteRow.Resource = NewResource

            UndoRedo.ItemPasted(NewResource, ParentResources)
        End If
    End Sub
#End Region

#Region "Paste verification sources"
    Private Sub PasteItems_VerificationSource(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)

        If PasteRow.Struct Is Nothing And strPasteColumnName.Contains("Struct") = False Then PasteRow.Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

        Dim selVerificationSource As VerificationSource = CType(selItem.Item, VerificationSource)
        Dim strSortNumber As String = PasteRow.VerificationSourceSort

        Select Case strPasteColumnName
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    Dim NewVerificationSource As New VerificationSource
                    Dim ParentVerificationSources As VerificationSources = PasteItems_VerificationSource_GetParentVerificationSources()
                    Dim intIndex As Integer = PasteItems_GetIndexOfVerificationSource(ParentVerificationSources)

                    Using copier As New ObjectCopy
                        NewVerificationSource = copier.CopyObject(selVerificationSource)
                    End Using

                    ParentVerificationSources.Insert(intIndex, NewVerificationSource)

                    UndoRedo.ItemPasted(NewVerificationSource, ParentVerificationSources)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "StructSort", "StructRTF"
                PasteItems_AsStruct(selItem, strPasteColumnName)
            Case "IndSort", "IndRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsIndicator(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "AsmSort", "AsmRTF"
                PasteItems_AsAssumption(selItem, strPasteColumnName)
        End Select
    End Sub

    Private Function PasteItems_VerificationSource_GetParentVerificationSources() As VerificationSources
        Dim ParentVerificationSources As VerificationSources = Nothing
        Dim intRowIndex As Integer = Grid.IndexOf(PasteRow)

        If PasteRow.VerificationSource IsNot Nothing Then
            ParentVerificationSources = Me.Logframe.GetParentCollection(PasteRow.VerificationSource)
        Else
            Dim selGridRow As LogframeRow

            For i = intRowIndex To 0 Step -1
                selGridRow = Me.Grid(i)
                If selGridRow.VerificationSource IsNot Nothing Then
                    ParentVerificationSources = Me.Logframe.GetParentCollection(selGridRow.VerificationSource)
                    Exit For
                ElseIf selGridRow.Indicator IsNot Nothing Then
                    ParentVerificationSources = selGridRow.Indicator.VerificationSources
                    Exit For
                ElseIf String.IsNullOrEmpty(selGridRow.StructSort) = False Then
                    Dim NewIndicator As New Indicator
                    selGridRow.Struct.Indicators.Add(NewIndicator)
                    ParentVerificationSources = NewIndicator.VerificationSources
                    Exit For
                End If
            Next
        End If

        If ParentVerificationSources Is Nothing Then
            Dim ParentStruct As Struct = Nothing
            Select Case PasteRow.Section
                Case LogframeObject.SectionTypes.Goal
                    Dim NewGoal As New Goal
                    Logframe.Goals.Add(NewGoal)
                    ParentStruct = NewGoal
                Case LogframeObject.SectionTypes.Purpose
                    Dim NewPurpose As New Purpose
                    Logframe.Purposes.Add(NewPurpose)
                    ParentStruct = NewPurpose
                Case LogframeObject.SectionTypes.Output
                    Dim NewOutput As New Output

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)
                    ParentPurpose.Outputs.Add(NewOutput)
                    ParentStruct = NewOutput
                Case LogframeObject.SectionTypes.Activity
                    Dim NewActivity As New Activity

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)

                    If ParentPurpose.Outputs.Count = 0 Then
                        Dim NewOutput As New Output
                        ParentPurpose.Outputs.Add(NewOutput)
                    End If

                    Dim ParentOutput As Output = ParentPurpose.Outputs(ParentPurpose.Outputs.Count - 1)
                    ParentOutput.Activities.Add(NewActivity)
                    ParentStruct = NewActivity
            End Select

            If ParentStruct IsNot Nothing Then
                Dim NewIndicator As New Indicator
                ParentStruct.Indicators.Add(NewIndicator)
                ParentVerificationSources = NewIndicator.VerificationSources
            End If
        End If

        Return ParentVerificationSources
    End Function

    Private Sub PasteItems_AsVerificationSource(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Dim strSortNumber As String = PasteRow.VerificationSourceSort
        Dim NewVerificationSource As New VerificationSource(CType(selItem.Item, LogframeObject).RTF)
        Dim ParentVerificationSources As VerificationSources = PasteItems_VerificationSource_GetParentVerificationSources()

        If ParentVerificationSources IsNot Nothing Then
            Dim intIndex As Integer = PasteItems_GetIndexOfVerificationSource(ParentVerificationSources)

            ParentVerificationSources.Insert(intIndex, NewVerificationSource)
            PasteRow.VerificationSource = NewVerificationSource

            UndoRedo.ItemPasted(NewVerificationSource, ParentVerificationSources)
        End If
    End Sub
#End Region

#Region "Paste assumptions"
    Private Sub PasteItems_Assumption(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)

        If PasteRow.Struct Is Nothing And strPasteColumnName.Contains("Struct") = False Then PasteRow.Struct = Me.Grid.GetPreviousStruct(Grid.IndexOf(PasteRow))

        Dim selAssumption As Assumption = CType(selItem.Item, Assumption)
        Dim strSortNumber As String = PasteRow.AssumptionSort

        Select Case strPasteColumnName
            Case "AsmSort", "AsmRTF"
                Dim NewAssumption As New Assumption
                Dim ParentAssumptions As Assumptions = PasteItems_Assumption_GetParentAssumptions()
                Dim intIndex As Integer = PasteItems_GetIndexOfAssumption(ParentAssumptions)

                Using copier As New ObjectCopy
                    NewAssumption = copier.CopyObject(selAssumption)
                End Using

                ParentAssumptions.Insert(intIndex, NewAssumption)
                PasteRow.Assumption = NewAssumption

                UndoRedo.ItemPasted(NewAssumption, ParentAssumptions)

                'Dim NewAssumption As New Assumption
                'Dim intIndex As Integer = PasteItems_GetIndexOfAssumption()

                'Using copier As New ObjectCopy
                '    NewAssumption = copier.CopyObject(selAssumption)
                'End Using

                'ParentAssumptions.Insert(intIndex, NewAssumption)

                'UndoRedo.ItemPasted(NewAssumption, ParentAssumptions)
            Case "StructSort", "StructRTF"
                PasteItems_AsStruct(selItem, strPasteColumnName)
            Case "IndSort", "IndRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsIndicator(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
            Case "VerSort", "VerRTF"
                If IsResourceBudgetRow(PasteRow) = False Then
                    PasteItems_AsVerificationSource(selItem, strPasteColumnName)
                Else
                    PasteItems_AsResource(selItem, strPasteColumnName)
                End If
        End Select
    End Sub

    Private Function PasteItems_Assumption_GetParentAssumptions() As Assumptions
        Dim ParentAssumptions As Assumptions = Nothing
        Dim intRowIndex As Integer = Grid.IndexOf(PasteRow)

        If PasteRow.Assumption IsNot Nothing Then
            ParentAssumptions = Me.Logframe.GetParentCollection(PasteRow.Assumption)
        Else
            Dim selGridRow As LogframeRow

            For i = intRowIndex To 0 Step -1
                selGridRow = Me.Grid(i)
                If selGridRow.Assumption IsNot Nothing Then
                    ParentAssumptions = Me.Logframe.GetParentCollection(selGridRow.Assumption)
                    Exit For
                ElseIf selGridRow.Struct IsNot Nothing Then
                    ParentAssumptions = selGridRow.Struct.Assumptions
                    Exit For
                End If
            Next
        End If

        If ParentAssumptions Is Nothing Then
            Select Case PasteRow.Section
                Case LogframeObject.SectionTypes.Goal
                    Dim NewGoal As New Goal
                    Logframe.Goals.Add(NewGoal)
                    ParentAssumptions = NewGoal.Assumptions
                Case LogframeObject.SectionTypes.Purpose
                    Dim NewPurpose As New Purpose
                    Logframe.Purposes.Add(NewPurpose)
                    ParentAssumptions = NewPurpose.Assumptions
                Case LogframeObject.SectionTypes.Output
                    Dim NewOutput As New Output

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)
                    ParentPurpose.Outputs.Add(NewOutput)
                    ParentAssumptions = NewOutput.Assumptions
                Case LogframeObject.SectionTypes.Activity
                    Dim NewActivity As New Activity

                    If Logframe.Purposes.Count = 0 Then
                        Dim NewPurpose As New Purpose
                        Logframe.Purposes.Add(NewPurpose)
                    End If

                    Dim ParentPurpose As Purpose = Logframe.Purposes(Logframe.Purposes.Count - 1)

                    If ParentPurpose.Outputs.Count = 0 Then
                        Dim NewOutput As New Output
                        ParentPurpose.Outputs.Add(NewOutput)
                    End If

                    Dim ParentOutput As Output = ParentPurpose.Outputs(ParentPurpose.Outputs.Count - 1)
                    ParentOutput.Activities.Add(NewActivity)
                    ParentAssumptions = NewActivity.Assumptions
            End Select
        End If

        Return ParentAssumptions
    End Function

    Private Sub PasteItems_AsAssumption(ByVal selItem As ClipboardItem, ByVal strPasteColumnName As String)
        Dim strSortNumber As String = PasteRow.AssumptionSort
        Dim NewAssumption As New Assumption(CType(selItem.Item, LogframeObject).RTF)
        Dim ParentAssumptions As Assumptions = PasteItems_Assumption_GetParentAssumptions()

        If ParentAssumptions IsNot Nothing Then
            Dim intIndex As Integer = PasteItems_GetIndexOfAssumption(ParentAssumptions)

            ParentAssumptions.Insert(intIndex, NewAssumption)
            PasteRow.Assumption = NewAssumption

            UndoRedo.ItemPasted(NewAssumption, ParentAssumptions)
        End If
    End Sub
#End Region

#Region "Select text"
    Public Sub SelectText(ByVal intTextSelectionIndex As Integer)
        Me.TextSelectionIndex = intTextSelectionIndex

        ResetAllImages()
        ReloadImages()
        Invalidate()
    End Sub

    Public Sub HighlightTextInCell(ByVal intMatchIndex As Integer, ByVal intMatchLength As Integer)
        If CurrentItem(False) Is Nothing Then Exit Sub

        If intMatchLength > 0 Then
            If CurrentCell.IsInEditMode = True Then EndEdit()

            BeginEdit(False)
            Dim ctl As RichTextEditingControlLogframe = EditingControl
            If ctl IsNot Nothing Then ctl.Select(intMatchIndex, intMatchLength)
        End If
    End Sub
#End Region
End Class
