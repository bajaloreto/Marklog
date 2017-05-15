Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Output
    Inherits Struct

    Private objParentPurposeGuid As Guid

    <ScriptIgnore()> _
    Public WithEvents Activities As New Activities

    <ScriptIgnore()> _
    Public WithEvents KeyMoments As New KeyMoments

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Private Sub Activities_ActivityAdded(ByVal sender As Object, ByVal e As ActivityAddedEventArgs) Handles Activities.ActivityAdded
        Dim selActivity As Activity = e.Activity

        selActivity.idParent = Me.idStruct
        selActivity.ParentOutputGuid = Me.Guid
    End Sub

    Private Sub KeyMoments_KeyMomentAdded(ByVal sender As Object, ByVal e As KeyMomentAddedEventArgs) Handles KeyMoments.KeyMomentAdded
        Dim selKeyMoment As KeyMoment = e.KeyMoment

        If selKeyMoment.ParentOutputGuid <> Me.Guid Then
            'selKeyMoment.idOutput = Me.idStruct
            selKeyMoment.ParentOutputGuid = Me.Guid
        End If
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
                For Each selActivity As Activity In Me.Activities
                    selActivity.idParent = value
                Next
            End If
        End Set
    End Property

    Public Overrides Property Section() As Integer
        Get
            Return 3
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property

    Public Property ParentPurposeGuid() As Guid
        Get
            Return objParentPurposeGuid
        End Get
        Set(ByVal value As Guid)
            objParentPurposeGuid = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemName() As String
        Get
            Return My.Settings.setStruct3sing
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Overloads Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return My.Settings.setStruct3
        End Get
    End Property
#End Region
End Class

Public Class Outputs
    Inherits Structs

    Public Event OutputAdded(ByVal sender As Object, ByVal e As OutputAddedEventArgs)

    Public Sub New()

    End Sub

    Public Overloads Sub Add(ByVal output As Output)
        If output IsNot Nothing Then
            List.Add(output)
            RaiseEvent OutputAdded(Me, New OutputAddedEventArgs(output))
        End If
    End Sub

    Public Overloads Sub Insert(ByVal index As Integer, ByVal output As Output)
        MyBase.Insert(index, output)
        RaiseEvent OutputAdded(Me, New OutputAddedEventArgs(output))
    End Sub

    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Output
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Output)
            End If
        End Get
    End Property

    Public Function GetOutputByGuid(ByVal objGuid As Guid) As Output
        Dim selOutput As Output = Nothing
        For Each objOutput As Output In Me.List
            If objOutput.Guid = objGuid Then
                selOutput = objOutput
                Exit For
            End If
        Next
        Return selOutput
    End Function

    Public Function GetKeyMomentByGuid(ByVal objGuid As Guid) As KeyMoment
        Dim selKeyMoment As KeyMoment = Nothing
        For Each selOutput As Output In Me.List
            selKeyMoment = selOutput.KeyMoments.GetKeyMomentByGuid(objGuid)
            If selKeyMoment IsNot Nothing Then Exit For
        Next
        Return selKeyMoment
    End Function
End Class

Public Class OutputAddedEventArgs
    Inherits EventArgs

    Public Property Output As Output

    Public Sub New(ByVal objOutput As Output)
        MyBase.New()

        Me.Output = objOutput
    End Sub
End Class

