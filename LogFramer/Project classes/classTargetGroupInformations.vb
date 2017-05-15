Imports System.Xml.Serialization

Public Class TargetGroupInformation
    Private intIdTargetGroupInformation, intIdTargetGroup As Integer
    Private strName As String
    Private lstPropertyValues As New List(Of String)
    Private intPropertyType As Integer
    Private intNrDecimals As Integer
    Private strUnit As String
    Private boolStandardItem As Boolean
    Private objGuid, objParentTargetGroupGuid As Guid

    Public CheckListOptions As New CheckListOptions

    Public Enum PropertyTypes As Integer
        Text = 0
        Number = 1
        YesNo = 2
        List = 3
        DateType = 4
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal name As String, ByVal type As Integer, ByVal nrdecimals As Integer, _
                   ByVal unit As String, Optional ByVal standarditem As Boolean = False)
        Me.Name = name
        Me.Type = type
        Me.NrDecimals = nrdecimals
        Me.Unit = unit
        Me.StandardItem = standarditem
    End Sub

    Public Property idTargetGroupInformation As Integer
        Get
            Return intIdTargetGroupInformation
        End Get
        Set(ByVal value As Integer)
            intIdTargetGroupInformation = value
        End Set
    End Property

    Public Property idTargetGroup As Integer
        Get
            Return intIdTargetGroup
        End Get
        Set(ByVal value As Integer)
            intIdTargetGroup = value
        End Set
    End Property

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
        End Set
    End Property

    Public Property ParentTargetGroupGuid() As Guid
        Get
            Return objParentTargetGroupGuid
        End Get
        Set(ByVal value As Guid)
            objParentTargetGroupGuid = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return strName
        End Get
        Set(ByVal value As String)
            strName = value
            If Me.StandardItem = True Then Me.StandardItem = False
        End Set
    End Property

    Public Property Type() As Integer
        Get
            Return intPropertyType
        End Get
        Set(ByVal value As Integer)
            intPropertyType = value
            If Me.StandardItem = True Then Me.StandardItem = False
        End Set
    End Property

    <XmlIgnore()> _
    Public ReadOnly Property TypeName() As String
        Get
            Return LIST_TargetGroupInformationTypes(Me.Type)
        End Get
    End Property

    Public Property NrDecimals() As Integer
        Get
            Return intNrDecimals
        End Get
        Set(ByVal value As Integer)
            intNrDecimals = value
            If Me.StandardItem = True Then Me.StandardItem = False
        End Set
    End Property

    Public Property Unit() As String
        Get
            Return strUnit
        End Get
        Set(ByVal value As String)
            strUnit = value
            If Me.StandardItem = True Then Me.StandardItem = False
        End Set
    End Property

    Public Property StandardItem() As Boolean
        Get
            Return boolStandardItem
        End Get
        Set(ByVal value As Boolean)
            boolStandardItem = value
        End Set
    End Property

    <XmlIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_TargetGroupInformation
        End Get
    End Property

    <XmlIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_TargetGroupInformations
        End Get
    End Property
End Class

Public Class TargetGroupInformations
    Inherits System.Collections.CollectionBase

    Public Event TargetGroupInformationAdded(ByVal sender As Object, ByVal e As TargetGroupInformationAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal targetgroupinformation As TargetGroupInformation)
        If targetgroupinformation IsNot Nothing Then
            List.Add(targetgroupinformation)
            RaiseEvent TargetGroupInformationAdded(Me, New TargetGroupInformationAddedEventArgs(targetgroupinformation))
        End If
    End Sub

    Public Function Contains(ByVal targetgroupinformation As TargetGroupInformation) As Boolean
        Return List.Contains(targetgroupinformation)
    End Function

    Public Function IndexOf(ByVal targetgroupinformation As TargetGroupInformation) As Integer
        Return List.IndexOf(targetgroupinformation)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal targetgroupinformation As TargetGroupInformation)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, targetgroupinformation.ItemName))
        ElseIf index = Count Then
            List.Add(targetgroupinformation)
            RaiseEvent TargetGroupInformationAdded(Me, New TargetGroupInformationAddedEventArgs(targetgroupinformation))
        Else
            List.Insert(index, targetgroupinformation)
            RaiseEvent TargetGroupInformationAdded(Me, New TargetGroupInformationAddedEventArgs(targetgroupinformation))
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As TargetGroupInformation
        Get
            Return CType(List.Item(index), TargetGroupInformation)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, TargetGroupInformation.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal targetgroupinformation As TargetGroupInformation)
        If List.Contains(targetgroupinformation) Then List.Remove(targetgroupinformation)
    End Sub

    Public Function VerifyIfNameExists(ByVal strName As String) As Integer
        Dim intCount As Integer
        For Each selTargetGroupInformation As TargetGroupInformation In Me.List
            If selTargetGroupInformation.Name.StartsWith(strName) Then intCount += 1
        Next

        Return intCount
    End Function

    Public Function GetTargetGroupInformationByGuid(ByVal objGuid As Guid) As TargetGroupInformation
        Dim selTargetGroupInformation As TargetGroupInformation = Nothing
        For Each objTargetGroupInformation As TargetGroupInformation In Me.List
            If objTargetGroupInformation.Guid = objGuid Then
                selTargetGroupInformation = objTargetGroupInformation
                Exit For
            End If
        Next
        Return selTargetGroupInformation
    End Function

    Public Sub SetDefaultInformations(ByVal intTargetGroupType As Integer)
        Me.Clear()
        Select Case intTargetGroupType
            Case TargetGroup.TargetGroupTypes.Individual, TargetGroup.TargetGroupTypes.Other
                Add(New TargetGroupInformation(LANG_Name, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Address, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_PoBox, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_TelephoneNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Email, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
            Case TargetGroup.TargetGroupTypes.Family, TargetGroup.TargetGroupTypes.ExtendedFamily
                Add(New TargetGroupInformation(LANG_Name, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Address, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_PoBox, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_TelephoneNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Email, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberMales, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberFemales, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberChildren, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
            Case TargetGroup.TargetGroupTypes.Community
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberMales, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberFemales, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NumberChildren, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
            Case TargetGroup.TargetGroupTypes.Association
                Add(New TargetGroupInformation(LANG_AssociationName, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Dim AddressInfo As New TargetGroupInformation(LANG_AddressOf, TargetGroupInformation.PropertyTypes.List, 0, String.Empty, True)
                AddressInfo.CheckListOptions.AddRange(LIST_OfficeTypes)
                Add(AddressInfo)
                Add(New TargetGroupInformation(LANG_Address, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_PoBox, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NameContact, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Function, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_TelephoneNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Email, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Website, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Dim StatuteInfo As New TargetGroupInformation(LANG_Statute, TargetGroupInformation.PropertyTypes.List, 0, String.Empty, True)
                StatuteInfo.CheckListOptions.AddRange(LIST_NonProfitTypes)
                Add(StatuteInfo)
                Add(New TargetGroupInformation(LANG_DateFoundation, TargetGroupInformation.PropertyTypes.DateType, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_RegistrationDate, TargetGroupInformation.PropertyTypes.DateType, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_RegistrationNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_MembersNumber, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_StaffNumber, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_MainObjectives, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_MainActivities, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
            Case TargetGroup.TargetGroupTypes.Enterprise
                Add(New TargetGroupInformation(LANG_CompanyName, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_CompanyType, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Address, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_PoBox, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NameContact, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Function, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_TelephoneNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Email, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Website, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_StaffNumber, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_MainActivities, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_GrossRevenu, TargetGroupInformation.PropertyTypes.Number, 2, "EUR", True))
            Case TargetGroup.TargetGroupTypes.Authority, TargetGroup.TargetGroupTypes.LocalAuthority
                Add(New TargetGroupInformation(LANG_StructureName, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Address, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_PoBox, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Town, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Country, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_NameContact, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Function, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_TelephoneNumber, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Email, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_Website, TargetGroupInformation.PropertyTypes.Text, 0, String.Empty, True))
                Add(New TargetGroupInformation(LANG_StaffNumber, TargetGroupInformation.PropertyTypes.Number, 0, String.Empty, True))
        End Select
    End Sub
End Class

Public Class TargetGroupInformationAddedEventArgs
    Inherits EventArgs

    Public Property TargetGroupInformation As TargetGroupInformation

    Public Sub New(ByVal objTargetGroupInformation As TargetGroupInformation)
        MyBase.New()

        Me.TargetGroupInformation = objTargetGroupInformation
    End Sub
End Class

Public Class ChecklistOption
    Private strOptionName As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal optionname As String)
        Me.OptionName = optionname
    End Sub

    Public Property OptionName As String
        Get
            Return strOptionName
        End Get
        Set(ByVal value As String)
            strOptionName = value
        End Set
    End Property

End Class

Public Class CheckListOptions
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal checklistoption As ChecklistOption)
        If checklistoption IsNot Nothing Then _
            List.Add(checklistoption)
    End Sub

    Public Sub AddRange(ByVal strRange As String())
        For Each strOption As String In strRange
            List.Add(New ChecklistOption(strOption))
        Next
    End Sub

    Public Function Contains(ByVal checklistoption As ChecklistOption) As Boolean
        Return List.Contains(checklistoption)
    End Function

    Public Function IndexOf(ByVal checklistoption As ChecklistOption) As Integer
        Return List.IndexOf(checklistoption)
    End Function

    Public Sub Insert(ByVal index As Integer, ByVal checklistoption As ChecklistOption)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, LANG_ChecklistOption))
        ElseIf index = Count Then
            List.Add(checklistoption)
        Else
            List.Insert(index, checklistoption)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ChecklistOption
        Get
            Return CType(List.Item(index), ChecklistOption)
        End Get
    End Property

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, LANG_ChecklistOption))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal checklistoption As ChecklistOption)
        If List.Contains(checklistoption) Then List.Remove(checklistoption)
    End Sub
End Class
