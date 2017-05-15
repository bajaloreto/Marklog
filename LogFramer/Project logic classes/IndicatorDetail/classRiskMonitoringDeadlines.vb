Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class RiskMonitoringDeadline
    Inherits TargetDeadline

    Public Sub New()

    End Sub

    Public Sub New(ByVal deadline As Date)
        Me.Relative = False
        Me.Deadline = deadline
    End Sub

    Public Sub New(ByVal guidreferencemoment As Guid, ByVal period As Single, ByVal periodunit As Integer)
        Me.Relative = True
        Me.GuidReferenceMoment = guidreferencemoment
        Me.Period = period
        Me.PeriodUnit = periodunit
    End Sub
End Class

Public Class RiskMonitoringDeadlines
    Inherits System.Collections.CollectionBase

    Public Event RiskMonitoringDeadlineAdded(ByVal sender As Object, ByVal e As RiskMonitoringDeadlineAddedEventArgs)

#Region "Properties"
    <XmlIgnore()> _
    <ScriptIgnore()> _
    Default Public ReadOnly Property Item(ByVal index As Integer) As RiskMonitoringDeadline
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), RiskMonitoringDeadline)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub Add(ByVal deadline As RiskMonitoringDeadline)
        List.Add(deadline)
        RaiseEvent RiskMonitoringDeadlineAdded(Me, New RiskMonitoringDeadlineAddedEventArgs(deadline))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal deadline As RiskMonitoringDeadline)
        List.Insert(intIndex, deadline)
        RaiseEvent RiskMonitoringDeadlineAdded(Me, New RiskMonitoringDeadlineAddedEventArgs(deadline))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal deadline As RiskMonitoringDeadline)
        If Me.List.Contains(deadline) Then
            Me.List.Remove(deadline)
        End If
    End Sub

    Public Function IndexOf(ByVal deadline As RiskMonitoringDeadline) As Integer
        Return List.IndexOf(deadline)
    End Function

    Public Function Contains(ByVal deadline As RiskMonitoringDeadline) As Boolean
        Return List.Contains(deadline)
    End Function

    Public Function FindDeadlineByDate(ByVal datDeadline As Date) As Boolean
        For Each selDeadline As RiskMonitoringDeadline In Me.List
            If selDeadline.Deadline = datDeadline Then Return True
        Next

        Return False
    End Function

    Public Function Sort() As RiskMonitoringDeadlines
        Dim sorter As System.Collections.IComparer = New RiskMonitoringDeadlineComparer(Of RiskMonitoringDeadline)
        Me.InnerList.Sort(sorter)

        Return Me
    End Function

    Public Function GetRiskMonitoringDeadlineByGuid(ByVal objGuid As Guid) As RiskMonitoringDeadline
        Dim selMonitoringDeadline As RiskMonitoringDeadline = Nothing

        For Each objMonitoringDeadline As RiskMonitoringDeadline In Me.List
            If objMonitoringDeadline.Guid = objGuid Then
                selMonitoringDeadline = objMonitoringDeadline
                Exit For
            End If
        Next
        Return selMonitoringDeadline
    End Function
#End Region
End Class

Public Class RiskMonitoringDeadlineAddedEventArgs
    Inherits EventArgs

    Public Property RiskMonitoringDeadline As RiskMonitoringDeadline

    Public Sub New(ByVal objMonitoringDeadline As RiskMonitoringDeadline)
        MyBase.New()

        Me.RiskMonitoringDeadline = objMonitoringDeadline
    End Sub
End Class

Public Class RiskMonitoringDeadlineComparer(Of RiskMonitoringDeadline)
    Implements System.Collections.IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        If x Is Nothing Then
            Throw New ArgumentNullException("x")
        End If
        If x Is Nothing Then
            Throw New ArgumentNullException("y")
        End If

        ' Get values
        Dim a As Date = x.ExactDeadline
        Dim b As Date = y.ExactDeadline

        ' Check for null first
        If a > Date.MinValue AndAlso b = Date.MinValue Then
            Return 1
        End If

        If a = Date.MinValue AndAlso b > Date.MinValue Then
            Return -1
        End If

        If a = Date.MinValue AndAlso b = Date.MinValue Then
            Return 0
        End If

        Return a.CompareTo(b)
    End Function
End Class
