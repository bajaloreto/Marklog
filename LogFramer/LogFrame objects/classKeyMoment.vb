Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class KeyMoment
    Private intIdKeyMoment As Integer
    Private datKeyMoment As Date
    Private strDescription As String
    Private objReferenceMoment As New Guid
    Private intType As Integer
    Private boolRelative As Boolean = True
    Private sngPeriod As Single
    Private intPeriodUnit As Integer
    Private intPeriodDirection As Integer
    Private objGuid, objParentOutputGuid As Guid

    Public Enum DirectionValues
        NotDefined = 0
        Before = 1
        After = 2
    End Enum

    Public Enum Types
        UserDefined = 0
        ProjectStart = 1
        ProjectEnd = 2
        ContractStart = 3
        ContractEnd = 4
    End Enum

#Region "Properties"
    Public Property idKeyMoment As Integer
        Get
            Return intIdKeyMoment
        End Get
        Set(ByVal value As Integer)
            intIdKeyMoment = value
        End Set
    End Property

    Public Property Relative() As Boolean
        Get
            Return boolRelative
        End Get
        Set(ByVal value As Boolean)
            boolRelative = value
            If boolRelative = True Then
                datKeyMoment = Nothing
            Else
                If datKeyMoment = Date.MinValue Then datKeyMoment = Now.Date
                sngPeriod = 0
                intPeriodUnit = 0
                intPeriodDirection = 0
            End If
        End Set
    End Property

    Public Property KeyMoment() As Date
        Get
            Return datKeyMoment
        End Get
        Set(ByVal value As Date)
            datKeyMoment = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ExactDateKeyMoment() As Date
        Get
            If Me.Relative = False Then
                Return Me.KeyMoment
            Else
                Dim RefDate As Date = CurrentLogFrame.GetReferenceDateByGuid(Me.GuidReferenceMoment)
                Dim selDate As Date = RefDate
                Dim sngPeriod As Single
                If Me.PeriodDirection = DirectionValues.After Then
                    sngPeriod = Me.Period
                ElseIf Me.PeriodDirection = DirectionValues.Before Then
                    sngPeriod = Me.Period * -1
                End If

                If RefDate = Date.MinValue And sngPeriod < 0 Then Return selDate

                Select Case PeriodUnit
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

    Public Property Description() As String
        Get
            Return strDescription
        End Get
        Set(ByVal value As String)
            strDescription = value
        End Set
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

    Public Property PeriodDirection() As Integer
        Get
            Return intPeriodDirection
        End Get
        Set(ByVal value As Integer)
            intPeriodDirection = value
        End Set
    End Property

    Public Property GuidReferenceMoment() As Guid
        Get
            Return objReferenceMoment
        End Get
        Set(ByVal value As Guid)
            objReferenceMoment = value
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return intType
        End Get
        Set(ByVal value As Integer)
            intType = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ParentOutputGuid() As Guid
        Get
            Return objParentOutputGuid
        End Get
        Set(ByVal value As Guid)
            objParentOutputGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_KeyMoment
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_KeyMoments
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal OutputGuid As Guid)
        Me.ParentOutputGuid = OutputGuid
    End Sub

    Public Sub New(ByVal keymoment As Date, ByVal description As String, Optional ByVal type As Integer = 0)
        Me.KeyMoment = keymoment
        Me.Description = description
        Me.Type = type
        Me.Relative = False
    End Sub

    Public Overrides Function ToString() As String
        Return ExactDateKeyMoment & ": " & Description
    End Function
#End Region
End Class

Public Class KeyMoments
    Inherits System.Collections.CollectionBase

    Public Event KeyMomentAdded(ByVal sender As Object, ByVal e As KeyMomentAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal keymoment As KeyMoment)
        List.Add(keymoment)

        RaiseEvent KeyMomentAdded(Me, New KeyMomentAddedEventArgs(keymoment))
    End Sub

    Public Sub AddRange(ByVal keymoments As KeyMoments)
        InnerList.AddRange(keymoments)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal keymoment As KeyMoment)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, keymoment.ItemName))
        ElseIf index = Count Then
            List.Add(keymoment)

            RaiseEvent KeyMomentAdded(Me, New KeyMomentAddedEventArgs(keymoment))
        Else
            List.Insert(index, keymoment)

            RaiseEvent KeyMomentAdded(Me, New KeyMomentAddedEventArgs(keymoment))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, KeyMoment.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal keymoment As KeyMoment)
        If List.Contains(keymoment) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, keymoment.ItemName))
        Else
            List.Remove(keymoment)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As KeyMoment
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), KeyMoment)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal keymoment As KeyMoment) As Integer
        Return List.IndexOf(keymoment)
    End Function

    Public Function Contains(ByVal keymoment As KeyMoment) As Boolean
        Return List.Contains(keymoment)
    End Function

    Public Function Sort() As KeyMoments
        Dim sorter As System.Collections.IComparer = New KeyMomentComparer(Of KeyMoment)
        Me.InnerList.Sort(sorter)

        Return Me
    End Function

    Public Function VerifyIfDescriptionExists(ByVal strDescription As String) As Integer
        Dim intCount As Integer
        For Each selKeyMoment As KeyMoment In Me.List
            If selKeyMoment.Description.StartsWith(strDescription) Then intCount += 1
        Next

        Return intCount
    End Function

    Public Function GetKeyMomentByGuid(ByVal objGuid As Guid) As KeyMoment
        Dim selKeyMoment As KeyMoment = Nothing
        For Each objKeyMoment As KeyMoment In Me.List
            If objKeyMoment.Guid = objGuid Then
                selKeyMoment = objKeyMoment
                Exit For
            End If
        Next
        Return selKeyMoment
    End Function
End Class

Public Class KeyMomentAddedEventArgs
    Inherits EventArgs

    Public Property KeyMoment As KeyMoment

    Public Sub New(ByVal objKeyMoment As KeyMoment)
        MyBase.New()

        Me.KeyMoment = objKeyMoment
    End Sub
End Class

Public Class KeyMomentComparer(Of KeyMoment)
    Implements System.Collections.IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        If x Is Nothing Then
            Throw New ArgumentNullException("x")
        End If
        If x Is Nothing Then
            Throw New ArgumentNullException("y")
        End If

        ' Get values
        Dim a As Date = x.ExactDateKeyMoment
        Dim b As Date = y.ExactDateKeyMoment

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


