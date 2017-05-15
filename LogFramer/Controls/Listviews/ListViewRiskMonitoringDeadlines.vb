Public Class ListViewRiskMonitoringDeadlines
    Inherits ListViewSortable

    Private objRiskMonitoring As RiskMonitoring
    Private strUnit, strValueName As String
    Private intNrDecimals As Integer
    Private intScoreSystem As Integer
    Private boolFromChildIndicator As Boolean

    Public Property RiskMonitoring() As RiskMonitoring
        Get
            Return objRiskMonitoring
        End Get
        Set(ByVal value As RiskMonitoring)
            objRiskMonitoring = value

            LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedRiskMonitoringDeadlines() As RiskMonitoringDeadline()
        Get
            Dim objRiskMonitoringDeadlines(SelectedItems.Count - 1) As RiskMonitoringDeadline

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objRiskMonitoringDeadlines(i) = RiskMonitoring.RiskMonitoringDeadlines.GetRiskMonitoringDeadlineByGuid(Me.SelectedGuids(i))
                Next

                Return objRiskMonitoringDeadlines
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Dates, 150, HorizontalAlignment.Center)
        End With
    End Sub

    Public Sub New(ByVal riskmonitoring As RiskMonitoring)
        Me.RiskMonitoring = riskmonitoring

        View = View.Details
        FullRowSelect = True
        Columns.Add(LANG_Dates, 150, HorizontalAlignment.Center)

        LoadItems()
    End Sub

    Public Sub LoadItems()

        Me.Items.Clear()
        If Me.RiskMonitoring IsNot Nothing Then
            For Each selRiskMonitoringDeadline As RiskMonitoringDeadline In RiskMonitoring.RiskMonitoringDeadlines.Sort
                Dim newItem As New ListViewItem(selRiskMonitoringDeadline.ExactDeadline.ToShortDateString)
                newItem.Name = selRiskMonitoringDeadline.Guid.ToString

                Me.Items.Add(newItem)
            Next
        End If
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
        EditItem()
        CurrentControl = Me
    End Sub

    Public Overrides Sub NewItem()
        PopUpRiskMonitoringDeadlineDetails(Nothing)
        CurrentControl = Me
    End Sub

    Public Overrides Sub EditItem()
        If RiskMonitoring.RiskMonitoringDeadlines.Count > 0 AndAlso Me.SelectedRiskMonitoringDeadlines.Length > 0 Then
            PopUpRiskMonitoringDeadlineDetails(Me.SelectedRiskMonitoringDeadlines(0))
            CurrentControl = Me
        End If
    End Sub

    Public Overrides Sub RemoveItem()
        If RiskMonitoring.RiskMonitoringDeadlines.Count > 0 AndAlso Me.SelectedRiskMonitoringDeadlines.Length > 0 Then
            If MsgBox(LANG_RemoveRiskMonitoringDeadline, MsgBoxStyle.OkCancel, LANG_Remove) = MsgBoxResult.Ok Then
                Dim selRiskMonitoringDeadline As RiskMonitoringDeadline = Me.SelectedRiskMonitoringDeadlines(0)
                ''CurrentUndoList.DeleteOperation(selRiskMonitoringDeadline, Me.RiskMonitoring.RiskMonitoringDeadlines, , , True)

                RiskMonitoring.RiskMonitoringDeadlines.Remove(selRiskMonitoringDeadline)
                Me.LoadItems()
            End If
        End If
    End Sub

    Private Sub PopUpRiskMonitoringDeadlineDetails(ByVal selRiskMonitoringDeadline As RiskMonitoringDeadline)
        Dim boolNew As Boolean

        If selRiskMonitoringDeadline Is Nothing Then
            boolNew = True
            selRiskMonitoringDeadline = New RiskMonitoringDeadline()
        End If

        Dim dialogRiskMonitoringDeadline As New DialogTargetDeadline(selRiskMonitoringDeadline.Deadline)

        With dialogRiskMonitoringDeadline
            If .ShowDialog = DialogResult.OK Then
                selRiskMonitoringDeadline.Deadline = .Deadline

                If boolNew = True Then
                    RiskMonitoring.RiskMonitoringDeadlines.Add(selRiskMonitoringDeadline)

                    ''CurrentUndoList.InsertNewOperation(selRiskMonitoringDeadline, RiskMonitoring.RiskMonitoringDeadlines, , , True)
                End If

                Me.LoadItems()
            End If
        End With
    End Sub

    Public Overrides Sub CutItems()
        CopyItems()

        For Each selRiskMonitoringDeadline As RiskMonitoringDeadline In SelectedRiskMonitoringDeadlines
            ''CurrentUndoList.CutOperation(selRiskMonitoringDeadline, RiskMonitoring.RiskMonitoringDeadlines, RiskMonitoring.RiskMonitoringDeadlines.IndexOf(selRiskMonitoringDeadline), , True)

            RiskMonitoring.RiskMonitoringDeadlines.Remove(selRiskMonitoringDeadline)
        Next

        LoadItems()
    End Sub

    Public Overrides Sub CopyItems()
        Dim CopyGroup As Date = Now()

        For Each selRiskMonitoringDeadline As RiskMonitoringDeadline In SelectedRiskMonitoringDeadlines
            Dim NewItem As New ClipboardItem(CopyGroup, selRiskMonitoringDeadline, RiskMonitoringDeadline.ItemName, 0)
            ItemClipboard.Insert(0, NewItem)
        Next
    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)
        Dim selItem As ClipboardItem
        Dim selRiskMonitoringDeadline As RiskMonitoringDeadline

        For i = 0 To PasteItems.Count - 1
            selItem = PasteItems(i)
            Select Case selItem.Item.GetType
                Case GetType(RiskMonitoringDeadline)
                    selRiskMonitoringDeadline = CType(selItem.Item, RiskMonitoringDeadline)
                    Dim NewRiskMonitoringDeadline As New RiskMonitoringDeadline

                    Using copier As New ObjectCopy
                        NewRiskMonitoringDeadline = copier.CopyObject(selRiskMonitoringDeadline)
                    End Using

                    RiskMonitoring.RiskMonitoringDeadlines.Add(NewRiskMonitoringDeadline)

                    ''CurrentUndoList.PasteOperation(NewRiskMonitoringDeadline, RiskMonitoring.RiskMonitoringDeadlines, 0, , True)
            End Select
        Next

        Me.LoadItems()
    End Sub
End Class

