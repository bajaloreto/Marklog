Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class VerificationSource
    Inherits LogframeObject

    Private intIdVerificationSource As Integer = -1
    Private intIdIndicator As Integer
    Private strResponsibility As String
    Private strFrequency As String
    Private strSource As String
    Private strWebsite As String
    Private strCollectionMethod As String
    Private objParentIndicatorGuid As Guid

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Sub New(ByVal Section As Integer, ByVal RTF As String)
        Me.Section = Section
        Me.RTF = RTF
    End Sub

#Region "Properties"
    Public Property idVerificationSource As Integer
        Get
            Return intIdVerificationSource
        End Get
        Set(ByVal value As Integer)
            intIdVerificationSource = value
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

    Public Property Responsibility() As String
        Get
            Return strResponsibility
        End Get
        Set(ByVal value As String)
            strResponsibility = value
        End Set
    End Property

    Public Property Frequency() As String
        Get
            Return strFrequency
        End Get
        Set(ByVal value As String)
            strFrequency = value
        End Set
    End Property

    Public Property Source() As String
        Get
            Return strSource
        End Get
        Set(ByVal value As String)
            strSource = value
        End Set
    End Property

    Public Property Website() As String
        Get
            Return strWebsite
        End Get
        Set(ByVal value As String)
            strWebsite = value
        End Set
    End Property

    Public Property CollectionMethod() As String
        Get
            Return strCollectionMethod
        End Get
        Set(ByVal value As String)
            strCollectionMethod = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_VerificationSource
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_VerificationSources
        End Get
    End Property
#End Region

#Region "Events"
    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        'no child collections
    End Sub
#End Region

End Class

Public Class VerificationSources
    Inherits System.Collections.CollectionBase

    Public Event VerificationSourceAdded(ByVal sender As Object, ByVal e As VerificationSourceAddedEventArgs)

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As VerificationSource
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), VerificationSource)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal verificationsource As VerificationSource)
        List.Add(verificationsource)
        RaiseEvent VerificationSourceAdded(Me, New VerificationSourceAddedEventArgs(verificationsource))
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal verificationsource As VerificationSource)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotInsert, verificationsource.ItemName))
        ElseIf index = Count Then
            List.Add(verificationsource)
            RaiseEvent VerificationSourceAdded(Me, New VerificationSourceAddedEventArgs(verificationsource))
        Else
            List.Insert(index, verificationsource)
            RaiseEvent VerificationSourceAdded(Me, New VerificationSourceAddedEventArgs(verificationsource))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, VerificationSource.ItemName))
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal verificationsource As VerificationSource)
        If List.Contains(verificationsource) = False Then
            System.Windows.Forms.MessageBox.Show(String.Format(ERR_IndexNotValidCannotRemove, verificationsource.ItemName))
        Else
            List.Remove(verificationsource)
        End If
    End Sub

    Public Function IndexOf(ByVal verificationsource As VerificationSource) As Integer
        Return List.IndexOf(verificationsource)
    End Function

    Public Function Contains(ByVal verificationsource As VerificationSource) As Boolean
        Return List.Contains(verificationsource)
    End Function
#End Region

#Region "Get by GUID"
    Public Function GetVerificationSourceByGuid(ByVal objGuid As Guid) As VerificationSource
        Dim selVerificationSource As VerificationSource = Nothing
        For Each objVerificationSource As VerificationSource In Me.List
            If objVerificationSource.Guid = objGuid Then
                selVerificationSource = objVerificationSource
                Exit For
            End If
        Next
        Return selVerificationSource
    End Function
#End Region

End Class

Public Class VerificationSourceAddedEventArgs
    Inherits EventArgs

    Public Property VerificationSource As VerificationSource

    Public Sub New(ByVal objVerificationSource As VerificationSource)
        MyBase.New()

        Me.VerificationSource = objVerificationSource
    End Sub
End Class
