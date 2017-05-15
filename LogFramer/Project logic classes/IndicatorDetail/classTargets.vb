Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Target
    Inherits ValueRange

    <ScriptIgnore()> _
    Public WithEvents BooleanValues As New BooleanValues 'one per indicator.responseclass

    <ScriptIgnore()> _
    Public WithEvents BooleanValuesMatrix As New BooleanValuesMatrix

    <ScriptIgnore()> _
    Public WithEvents DoubleValues As New DoubleValues

    <ScriptIgnore()> _
    Public WithEvents DoubleValuesMatrix As New DoubleValuesMatrix

    <ScriptIgnore()> _
    Public WithEvents AudioVisualDetail As New AudioVisualDetail

    Private intIdTarget As Integer = -1
    Private intIdIndicator As Integer
    Private objTargetDeadlineGuid, objParentIndicatorGuid, objReferenceTargetGuid As Guid
    Private boolRelative As Boolean
    Private strTargetName As String
    Private strFormula As String
    Private dlbScore As Double

    Public Sub New()

    End Sub

    Public Sub New(ByVal targetdeadlineguid As Guid)
        Me.TargetDeadlineGuid = targetdeadlineguid
    End Sub

    Public Sub New(ByVal targetdeadlineguid As Guid, ByVal opmin As String, ByVal minvalue As Double, ByVal opmax As String, ByVal maxvalue As Double)
        Me.TargetDeadlineGuid = targetdeadlineguid
        Me.OpMin = opmin
        Me.MinValue = minvalue
        Me.OpMax = opmax
        Me.MaxValue = maxvalue
    End Sub

    Public Sub New(ByVal targetdeadlineguid As Guid, ByVal scorevalue As Double)
        Me.TargetDeadlineGuid = targetdeadlineguid
        Me.Score = scorevalue
    End Sub

    Public Property idTarget As Integer
        Get
            Return intIdTarget
        End Get
        Set(ByVal value As Integer)
            intIdTarget = value
        End Set
    End Property

    Public Property idIndicator As Integer
        Get
            Return intIdIndicator
        End Get
        Set(ByVal value As Integer)
            intIdIndicator = value
        End Set
    End Property

    Public Property ParentIndicatorGuid() As Guid
        Get
            Return objParentIndicatorGuid
        End Get
        Set(ByVal value As Guid)
            objParentIndicatorGuid = value
        End Set
    End Property

    Public Property TargetDeadlineGuid As Guid
        Get
            Return objTargetDeadlineGuid
        End Get
        Set(ByVal value As Guid)
            objTargetDeadlineGuid = value
        End Set
    End Property

    Public Property ReferenceTargetGuid As Guid
        Get
            Return objReferenceTargetGuid
        End Get
        Set(ByVal value As Guid)
            objReferenceTargetGuid = value
        End Set
    End Property

    Public Property TargetName As String
        Get
            Return strTargetName
        End Get
        Set(ByVal value As String)
            strTargetName = value
        End Set
    End Property

    Public Property Score As Double
        Get
            Return dlbScore
        End Get
        Set(ByVal value As Double)
            dlbScore = value
        End Set
    End Property

    Public Property Relative() As Boolean
        Get
            Return boolRelative
        End Get
        Set(ByVal value As Boolean)
            boolRelative = value
        End Set
    End Property

    Public Property Formula As String
        Get
            Return strFormula
        End Get
        Set(value As String)
            strFormula = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Target
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Targets
        End Get
    End Property

    Public Function FormatTarget(ByVal strValueName As String, ByVal intNrDecimals As Integer, ByVal strUnit As String) As String
        Dim strRange As String = String.Empty, strOp As String

        If String.IsNullOrEmpty(OpMin) = False And String.IsNullOrEmpty(OpMax) = False Then
            If OpMin = CONST_Equals Then strOp = CONST_Equals Else strOp = CONST_SmallerThan
            If Len(OpMin) > 1 Then strOp &= CONST_Equals
            strRange = DisplayAsUnit(MinValue, intNrDecimals, strUnit) & " " & strOp & strValueName & _
                OpMax & " " & DisplayAsUnit(MaxValue, intNrDecimals, strUnit)
        Else
            If String.IsNullOrEmpty(OpMin) = False Then
                strRange = strValueName & " " & OpMin & DisplayAsUnit(MinValue, intNrDecimals, strUnit)
            ElseIf String.IsNullOrEmpty(OpMax) = False Then
                strRange = strValueName & " " & OpMax & DisplayAsUnit(MaxValue, intNrDecimals, strUnit)
            Else
                strRange = "..."
            End If
        End If
        Return strRange
    End Function

    Private Sub DoubleValues_DoubleValueAdded(ByVal sender As Object, ByVal e As DoubleValueAddedEventArgs) Handles DoubleValues.DoubleValueAdded
        Dim selDoubleValue As DoubleValue = e.DoubleValue

        selDoubleValue.idParent = Me.idTarget
        selDoubleValue.ParentGuid = Me.Guid
    End Sub

    Private Sub BooleanValues_BooleanValueAdded(sender As Object, e As BooleanValueAddedEventArgs) Handles BooleanValues.BooleanValueAdded
        Dim selBooleanValue As BooleanValue = e.BooleanValue

        selBooleanValue.idParent = Me.idTarget
        selBooleanValue.ParentGuid = Me.Guid
    End Sub
End Class

Public Class Targets
    Inherits System.Collections.CollectionBase

    Public Event TargetAdded(ByVal sender As Object, ByVal e As TargetAddedEventArgs)

    Public Sub Add(ByVal target As Target)
        List.Add(target)
        RaiseEvent TargetAdded(Me, New TargetAddedEventArgs(target))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal target As Target)
        List.Insert(intIndex, target)
        RaiseEvent TargetAdded(Me, New TargetAddedEventArgs(target))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal target As Target)
        If Me.List.Contains(target) Then
            Me.List.Remove(target)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Target
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), Target)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal target As Target) As Integer
        Return List.IndexOf(target)
    End Function

    Public Function Contains(ByVal target As Target) As Boolean
        Return List.Contains(target)
    End Function

    'Public Function GetExactTargetMinValue(ByVal selTarget As Target) As Double
    '    Dim dblValue As Double

    '    If selTarget.Relative = False Then
    '        dblValue = selTarget.MinValue
    '    Else
    '        Dim RefTarget As Target = GetTargetByGuid(selTarget.ReferenceTargetGuid)

    '        If RefTarget IsNot Nothing Then
    '            Dim Eparser As New ExpressionParser


    '            For Each objTarget As Target In Me.List
    '                If objTarget.Guid <> selTarget.Guid Then
    '                    Eparser.VariableList.Add(objTarget.TargetName, objTarget.MinValue)
    '                End If
    '            Next
    '            Eparser.Expression = RefTarget.Formula
    '            dblValue = Eparser.Parse

    '        End If
    '    End If

    '    Return dblValue
    'End Function

    Public Sub ConformWithTargets(ByVal selResponseValues As ResponseValues)
        Dim selTarget As Target
        For i = 0 To Me.Count - 1
            selTarget = Me(i)
            If i > selResponseValues.Count - 1 Then
                selResponseValues.Add(New ResponseValue(selTarget.Guid, 0))
            Else
                If selResponseValues(i).TargetGuid <> selTarget.Guid Then
                    selResponseValues(i).TargetGuid = selTarget.Guid
                End If
            End If
        Next

        If selResponseValues.Count > Me.Count Then
            For i = selResponseValues.Count - 1 To Me.Count Step -1
                selResponseValues.Remove(i)
            Next
        End If
    End Sub

    Public Function GetHighestScoreOfTargets() As Double
        Dim dblHighestScore As Double

        For Each selTarget As Target In Me.List
            If selTarget.Score > dblHighestScore Then dblHighestScore = selTarget.Score
        Next

        Return dblHighestScore
    End Function

    Public Function GetTargetByGuid(ByVal objGuid As Guid) As Target
        Dim selTarget As Target = Nothing

        For Each objTarget As Target In Me.List
            If objTarget.Guid = objGuid Then
                selTarget = objTarget
                Exit For
            End If
        Next
        Return selTarget
    End Function
End Class

Public Class TargetAddedEventArgs
    Inherits EventArgs

    Public Property Target As Target

    Public Sub New(ByVal objTarget As Target)
        MyBase.New()

        Me.Target = objTarget
    End Sub
End Class

Public Class TargetUpdatedEventArgs
    Inherits EventArgs

    Public Property TargetIndex As Integer

    Public Sub New(ByVal intTargetIndex As Integer)
        MyBase.New()

        Me.TargetIndex = intTargetIndex
    End Sub
End Class

'Public Class TargetComparer(Of Target)
'    Implements System.Collections.IComparer

'    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
'        If x Is Nothing Then
'            Throw New ArgumentNullException("x")
'        End If
'        If x Is Nothing Then
'            Throw New ArgumentNullException("y")
'        End If

'        ' Get values
'        Dim a As Date = x.Deadline
'        Dim b As Date = y.Deadline

'        ' Check for null first
'        If a > Date.MinValue AndAlso b = Date.MinValue Then
'            Return 1
'        End If

'        If a = Date.MinValue AndAlso b > Date.MinValue Then
'            Return -1
'        End If

'        If a = Date.MinValue AndAlso b = Date.MinValue Then
'            Return 0
'        End If

'        Return a.CompareTo(b)
'    End Function
'End Class
