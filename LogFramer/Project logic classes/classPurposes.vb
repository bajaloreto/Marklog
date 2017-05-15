Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Purpose
    Inherits Struct

    <ScriptIgnore()> _
    Public WithEvents Outputs As New Outputs

    <ScriptIgnore()> _
    Public WithEvents TargetGroups As New TargetGroups

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Private Sub Outputs_OutputAdded(ByVal sender As Object, ByVal e As OutputAddedEventArgs) Handles Outputs.OutputAdded
        Dim selOutput As Output = e.Output

        selOutput.idParent = Me.idStruct
        selOutput.ParentPurposeGuid = Me.Guid
    End Sub

    Private Sub TargetGroups_TargetGroupAdded(ByVal sender As Object, ByVal e As TargetGroupAddedEventArgs) Handles TargetGroups.TargetGroupAdded
        Dim selTargetGroup As TargetGroup = e.TargetGroup

        'selTargetGroup.idPurpose = Me.idStruct
        selTargetGroup.ParentPurposeGuid = Me.Guid
    End Sub
#End Region

#Region "Properties"
    Public Overrides Property idStruct As Integer
        Get
            Return MyBase.idStruct
        End Get
        Set(ByVal value As Integer)
            MyBase.idStruct = value

            If value > 0 Then
                For Each selOutput As Output In Me.Outputs
                    selOutput.idParent = value
                Next
            End If
        End Set
    End Property

    Public Overrides Property Section() As Integer
        Get
            Return 2
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return My.Settings.setStruct2sing
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return My.Settings.setStruct2
        End Get
    End Property
#End Region
End Class

Public Class Purposes
    Inherits Structs

    Public Sub New()

    End Sub

    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Purpose
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Purpose)
            End If
        End Get
    End Property

    Public Function GetPurposeByGuid(ByVal objGuid As Guid) As Purpose
        Dim selPurpose As Purpose = Nothing
        For Each objPurpose As Purpose In Me.List
            If objPurpose.Guid = objGuid Then
                selPurpose = objPurpose
                Exit For
            End If
        Next
        Return selPurpose
    End Function

    Public Function GetTargetGroupByGuid(ByVal objGuid As Guid) As TargetGroup
        Dim selTargetGroup As TargetGroup = Nothing
        For Each selPurpose As Purpose In Me.List
            selTargetGroup = selPurpose.TargetGroups.GetTargetGroupByGuid(objGuid)
            If selTargetGroup IsNot Nothing Then Exit For
        Next
        Return selTargetGroup
    End Function
End Class
