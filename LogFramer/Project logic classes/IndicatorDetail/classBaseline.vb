Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Baseline
    Inherits LogframeObject

    <ScriptIgnore()> _
    Public WithEvents BooleanValues As New BooleanValues 'one per indicator.responseclass/statement

    <ScriptIgnore()> _
    Public WithEvents BooleanValuesMatrix As New BooleanValuesMatrix

    <ScriptIgnore()> _
    Public WithEvents DoubleValues As New DoubleValues

    <ScriptIgnore()> _
    Public WithEvents DoubleValuesMatrix As New DoubleValuesMatrix

    <ScriptIgnore()> _
    Public WithEvents AudioVisualDetail As New AudioVisualDetail

    Private intIdBaseline As Integer = -1
    Private objParentIndicatorGuid As Guid
    Private dblValue As Double, dblScore, dblPopulationPercentage As Double

    Public Sub New()

    End Sub

    Public Sub New(ByVal score As Double)
        Me.Score = score
    End Sub

#Region "Properties"
    Public Property idBaseline As Integer
        Get
            Return intIdBaseline
        End Get
        Set(ByVal value As Integer)
            intIdBaseline = value
        End Set
    End Property

    Public Property ParentIndicatorGuid As Guid
        Get
            Return objParentIndicatorGuid
        End Get
        Set(ByVal value As Guid)
            objParentIndicatorGuid = value
        End Set
    End Property

    Public Property Score As Double
        Get
            Return dblScore
        End Get
        Set(ByVal value As Double)
            dblScore = value
        End Set
    End Property

    Public Property Value As Double
        Get
            Return dblValue
        End Get
        Set(ByVal value As Double)
            dblValue = value
        End Set
    End Property

    Public Property PopulationPercentage As Double
        Get
            Return dblPopulationPercentage
        End Get
        Set(ByVal value As Double)
            dblPopulationPercentage = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Baseline
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Baseline
        End Get
    End Property
#End Region

#Region "Events"
    Private Sub DoubleValues_DoubleValueAdded(ByVal sender As Object, ByVal e As DoubleValueAddedEventArgs) Handles DoubleValues.DoubleValueAdded
        Dim selDoubleValue As DoubleValue = e.DoubleValue

        selDoubleValue.idParent = Me.idBaseline
        selDoubleValue.ParentGuid = Me.Guid
    End Sub

    Private Sub BooleanValues_BooleanValueAdded(sender As Object, e As BooleanValueAddedEventArgs) Handles BooleanValues.BooleanValueAdded
        Dim selBooleanValue As BooleanValue = e.BooleanValue

        selBooleanValue.idParent = Me.idBaseline
        selBooleanValue.ParentGuid = Me.Guid
    End Sub

    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selDoubleValue As DoubleValue In Me.DoubleValues
            selDoubleValue.ParentGuid = Me.Guid
        Next
        For Each selBooleanValue As BooleanValue In Me.BooleanValues
            selBooleanValue.ParentGuid = Me.Guid
        Next
    End Sub
#End Region
End Class


