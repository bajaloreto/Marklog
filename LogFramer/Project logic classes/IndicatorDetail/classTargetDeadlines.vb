Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class TargetDeadline
    Private objGuid, objGuidReferenceMoment As Guid
    Private datDeadline As Date
    Private boolRelative As Boolean
    Private sngPeriod As Single
    Private intPeriodUnit As Integer

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

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property Relative() As Boolean
        Get
            Return boolRelative
        End Get
        Set(ByVal value As Boolean)
            boolRelative = value
            If boolRelative = True Then
                datDeadline = Nothing
            Else
                If datDeadline = Date.MinValue Then datDeadline = Now.Date
                sngPeriod = 0
                intPeriodUnit = 0
            End If
        End Set
    End Property

    Public Property Deadline() As Date
        Get
            Return datDeadline
        End Get
        Set(ByVal value As Date)
            datDeadline = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactDeadline() As Date
        Get
            If Me.Relative = False Then
                Return Me.Deadline
            Else

                Dim ProjectKeyMoment As KeyMoment = CurrentLogFrame.GetProjectStartKeyMoment
                If ProjectKeyMoment.Guid <> Me.GuidReferenceMoment Then ProjectKeyMoment = CurrentLogFrame.GetProjectEndKeyMoment
                If ProjectKeyMoment.Guid <> Me.GuidReferenceMoment Then Return Date.MinValue

                Dim RefDate As Date = ProjectKeyMoment.ExactDateKeyMoment
                Dim selDate As Date = RefDate
                Dim sngPeriod As Single = Me.Period


                If RefDate = Date.MinValue And sngPeriod < 0 Then Return selDate

                Select Case Me.PeriodUnit
                    Case DurationUnits.Day
                        selDate = RefDate.AddDays(sngPeriod)
                    Case DurationUnits.Hour
                        selDate = RefDate.AddHours(sngPeriod)
                    Case DurationUnits.Minute
                        selDate = RefDate.AddMinutes(sngPeriod)
                    Case DurationUnits.Month
                        selDate = RefDate.AddMonths(sngPeriod)
                    Case DurationUnits.Semester
                        selDate = RefDate.AddMonths(sngPeriod * 6)
                    Case DurationUnits.Trimester
                        selDate = RefDate.AddMonths(sngPeriod * 3)
                    Case DurationUnits.Week
                        selDate = RefDate.AddDays(sngPeriod * 7)
                    Case DurationUnits.Year
                        selDate = RefDate.AddYears(sngPeriod)
                End Select

                Return selDate
            End If
        End Get
    End Property

    Public Property Period() As Single
        Get
            Return sngPeriod
        End Get
        Set(ByVal value As Single)
            sngPeriod = value
        End Set
    End Property

    Public Property PeriodUnit() As Integer
        Get
            Return intPeriodUnit
        End Get
        Set(ByVal value As Integer)
            intPeriodUnit = value
        End Set
    End Property

    Public ReadOnly Property PeriodDirection() As Integer
        Get
            Return 2 'Directionvalues.After
        End Get
    End Property

    Public Property GuidReferenceMoment() As Guid
        Get
            Return objGuidReferenceMoment
        End Get
        Set(ByVal value As Guid)
            objGuidReferenceMoment = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_TargetDeadline
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_TargetDeadlines
        End Get
    End Property
End Class

Public Class TargetDeadlines
    Inherits System.Collections.CollectionBase

    Public Event TargetDeadlineAdded(ByVal sender As Object, ByVal e As TargetDeadlineAddedEventArgs)

#Region "Properties"
    <XmlIgnore()> _
    <ScriptIgnore()> _
    Default Public ReadOnly Property Item(ByVal index As Integer) As TargetDeadline
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), TargetDeadline)
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub ConformWithDeadlines(ByVal selTargets As Targets)
        Dim selTargetDeadline As TargetDeadline
        For i = 0 To Me.Count - 1
            selTargetDeadline = Me(i)
            If i > selTargets.Count - 1 Then
                selTargets.Add(New Target(selTargetDeadline.Guid, "", 0, "", 0))
            Else
                If selTargets(i).TargetDeadlineGuid <> selTargetDeadline.Guid Then
                    selTargets(i).TargetDeadlineGuid = selTargetDeadline.Guid
                End If
            End If
        Next

        If selTargets.Count > Me.Count Then
            For i = selTargets.Count - 1 To Me.Count Step -1
                selTargets.Remove(i)
            Next
        End If
    End Sub

    Public Sub ConformWithDeadlines(ByVal selPopulationTargets As PopulationTargets)
        Dim selTargetDeadline As TargetDeadline
        For i = 0 To Me.Count - 1
            selTargetDeadline = Me(i)
            If i > selPopulationTargets.Count - 1 Then
                selPopulationTargets.Add(New PopulationTarget(selTargetDeadline.Guid, 0))
            Else
                If selPopulationTargets(i).TargetDeadlineGuid <> selTargetDeadline.Guid Then
                    selPopulationTargets(i).TargetDeadlineGuid = selTargetDeadline.Guid
                End If
            End If
        Next

        If selPopulationTargets.Count > Me.Count Then
            For i = selPopulationTargets.Count - 1 To Me.Count Step -1
                selPopulationTargets.Remove(i)
            Next
        End If
    End Sub

    Public Sub Add(ByVal deadline As TargetDeadline)
        List.Add(deadline)
        RaiseEvent TargetDeadlineAdded(Me, New TargetDeadlineAddedEventArgs(deadline))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal deadline As TargetDeadline)
        List.Insert(intIndex, deadline)
        RaiseEvent TargetDeadlineAdded(Me, New TargetDeadlineAddedEventArgs(deadline))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal deadline As TargetDeadline)
        If Me.List.Contains(deadline) Then
            Me.List.Remove(deadline)
        End If
    End Sub

    Public Function IndexOf(ByVal deadline As TargetDeadline) As Integer
        Return List.IndexOf(deadline)
    End Function

    Public Function Contains(ByVal deadline As TargetDeadline) As Boolean
        Return List.Contains(deadline)
    End Function

    Public Function FindDeadlineByDate(ByVal datDeadline As Date) As Boolean
        For Each selDeadline As TargetDeadline In Me.List
            If selDeadline.Deadline = datDeadline Then Return True
        Next

        Return False
    End Function

    Public Function Sort() As TargetDeadlines
        Dim sorter As System.Collections.IComparer = New TargetDeadlineComparer(Of TargetDeadline)
        Me.InnerList.Sort(sorter)

        Return Me
    End Function

    Public Function GetTargetDeadlineByGuid(ByVal objGuid As Guid) As TargetDeadline
        Dim selTargetDeadline As TargetDeadline = Nothing

        For Each objTargetDeadline As TargetDeadline In Me.List
            If objTargetDeadline.Guid = objGuid Then
                selTargetDeadline = objTargetDeadline
                Exit For
            End If
        Next
        Return selTargetDeadline
    End Function
#End Region
End Class

Public Class TargetDeadlineAddedEventArgs
    Inherits EventArgs

    Public Property TargetDeadline As TargetDeadline

    Public Sub New(ByVal objTargetDeadline As TargetDeadline)
        MyBase.New()

        Me.TargetDeadline = objTargetDeadline
    End Sub
End Class

Public Class TargetDeadlineComparer(Of TargetDeadline)
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
