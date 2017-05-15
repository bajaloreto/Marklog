Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class TargetGroup
    Public Event GuidChanged(ByVal sender As Object)

    Private intIdTargetGroup As Integer
    Private strName As String
    Private intType As Integer
    Private intNumber As Integer
    Private intNumberOfMales, intNumberOfFemales As Integer
    Private intNumberOfPeople As Integer
    Private strLocation As String
    Private objGuid, objParentPurposeGuid As Guid

    <ScriptIgnore()> _
    Public WithEvents TargetGroupInformations As New TargetGroupInformations

    Public Enum TargetGroupTypes As Integer
        Individual = 0
        Family = 1
        ExtendedFamily = 2
        Community = 3
        Association = 4
        Enterprise = 5
        LocalAuthority = 6
        Authority = 7
        Other = 8
    End Enum

#Region "Properties"
    Public Property idTargetGroup As Integer
        Get
            Return intIdTargetGroup
        End Get
        Set(ByVal value As Integer)
            intIdTargetGroup = value

            If value > 0 Then
                For Each selTargetGroupInfo As TargetGroupInformation In Me.TargetGroupInformations
                    selTargetGroupInfo.idTargetGroup = value
                Next
            End If
        End Set
    End Property

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
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

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property TypeName() As String
        Get
            Return LIST_TargetGroupTypes(Me.Type)
        End Get
    End Property

    Public Property Number() As Integer
        Get
            Return intNumber
        End Get
        Set(ByVal value As Integer)
            intNumber = value
        End Set
    End Property

    Public Property NumberOfMales() As Integer
        Get
            Return intNumberOfMales
        End Get
        Set(ByVal value As Integer)
            intNumberOfMales = value
        End Set
    End Property

    Public ReadOnly Property PercentageMales() As Single
        Get
            If intNumberOfMales <> 0 Then
                Return intNumberOfMales / intNumberOfPeople
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property NumberOfFemales() As Integer
        Get
            Return intNumberOfFemales
        End Get
        Set(ByVal value As Integer)
            intNumberOfFemales = value
        End Set
    End Property

    Public ReadOnly Property PercentageFemales() As Single
        Get
            If intNumberOfFemales <> 0 Then
                Return intNumberOfFemales / intNumberOfPeople
            Else
                Return 0
            End If
        End Get
    End Property

    Public Property NumberOfPeople() As Integer
        Get
            Return intNumberOfPeople
        End Get
        Set(ByVal value As Integer)
            intNumberOfPeople = value
        End Set
    End Property

    Public Property Location() As String
        Get
            Return strLocation
        End Get
        Set(ByVal value As String)
            strLocation = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then
                objGuid = Guid.NewGuid
                RaiseEvent GuidChanged(Me)
            End If
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
            RaiseEvent GuidChanged(Me)
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
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_TargetGroup
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_TargetGroups
        End Get
    End Property
#End Region

#Region "Methods and events"
    Public Sub New()

    End Sub

    Public Sub New(ByVal PurposeGuid As Guid)
        Me.ParentPurposeGuid = PurposeGuid
    End Sub

    Public Overrides Function ToString() As String
        Return Me.Name
    End Function

    Private Sub TargetGroupInformations_TargetGroupInformationAdded(ByVal sender As Object, ByVal e As TargetGroupInformationAddedEventArgs) Handles TargetGroupInformations.TargetGroupInformationAdded
        Dim selTargetGroupInformation As TargetGroupInformation = e.TargetGroupInformation

        selTargetGroupInformation.idTargetGroup = Me.idTargetGroup
        selTargetGroupInformation.ParentTargetGroupGuid = Me.Guid
    End Sub

    Private Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selTargetGroupInfo As TargetGroupInformation In Me.TargetGroupInformations
            selTargetGroupInfo.ParentTargetGroupGuid = Me.Guid
        Next
    End Sub
#End Region
End Class

Public Class TargetGroups
    Inherits System.Collections.CollectionBase

    Public Event TargetGroupAdded(ByVal sender As Object, ByVal e As TargetGroupAddedEventArgs)

    Public Sub Add(ByVal targetgroup As TargetGroup)
        List.Add(targetgroup)

        RaiseEvent TargetGroupAdded(Me, New TargetGroupAddedEventArgs(targetgroup))
    End Sub

    Public Sub AddRange(ByVal targetgroups As TargetGroups)
        InnerList.AddRange(targetgroups)
    End Sub

    Public Function Contains(ByVal targetgroup As TargetGroup) As Boolean
        Return List.Contains(targetgroup)
    End Function

    Public Function IndexOf(ByVal targetgroup As TargetGroup) As Integer
        Return List.IndexOf(targetgroup)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal targetgroup As TargetGroup)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, targetgroup.ItemName))
        ElseIf index = Count Then
            List.Add(targetgroup)

            RaiseEvent TargetGroupAdded(Me, New TargetGroupAddedEventArgs(targetgroup))
        Else
            List.Insert(index, targetgroup)

            RaiseEvent TargetGroupAdded(Me, New TargetGroupAddedEventArgs(targetgroup))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As TargetGroup
        Get
            Return CType(List.Item(index), TargetGroup)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, TargetGroup.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal targetgroup As TargetGroup)
        If List.Contains(targetgroup) Then List.Remove(targetgroup)
    End Sub

    Public Function VerifyIfNameExists(ByVal strName As String) As Integer
        Dim intCount As Integer
        For Each selTargetGroup As TargetGroup In Me.List
            If selTargetGroup.Name.StartsWith(strName) Then intCount += 1
        Next

        Return intCount
    End Function

    Public Function GetTargetGroupByGuid(ByVal objGuid As Guid) As TargetGroup
        Dim selTargetGroup As TargetGroup = Nothing
        For Each objTargetGroup As TargetGroup In Me.List
            If objTargetGroup.Guid = objGuid Then
                selTargetGroup = objTargetGroup
                Exit For
            End If
        Next
        Return selTargetGroup
    End Function
End Class

Public Class TargetGroupAddedEventArgs
    Inherits EventArgs

    Public Property TargetGroup As TargetGroup

    Public Sub New(ByVal objTargetGroup As TargetGroup)
        MyBase.New()

        Me.TargetGroup = objTargetGroup
    End Sub
End Class
