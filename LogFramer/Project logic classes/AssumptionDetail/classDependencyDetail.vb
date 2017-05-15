Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class DependencyDetail

    Private intDependencyType As Integer
    Private intInputType As Integer
    Private intImportanceLevel As Integer
    Private intDeliverableType As Integer
    Private strDeliverables As String
    Private strSupplier As String
    Private datDateExpected As Date
    Private datDateDelivered As Date

    Public Sub New()

    End Sub

#Region "Enumerations"
    Public Enum DependencyTypes As Integer
        'current programme
        InputToProgramme = 0
        Contextual = 1
        Climatological = 2

        'other programmes
        PrecedingProgramme = 10
        FollowingProgramme = 11
        ComplementaryProgrammeInput = 12
        ComplementaryProgrammeOutput = 13
        Other = 20
    End Enum

    Public Enum InputTypes As Integer
        'available inputs
        NotSet = 0
        KnowledgeExpertise = 1
        ManPower = 2
        CoFinancing = 3
        ServiceDelivery = 4
        GoodsDelivery = 5
        Logistics = 6
        Networking = 7
        Infrastructure = 8

        'approval of authorities
        AdministrativeApproval = 20
        DonorApproval = 21
        PoliticalDecision = 22
        AuditorsApproval = 23

        Other = 50
    End Enum

    Public Enum ImportanceLevels As Integer
        Low = 0
        Medium = 1
        High = 2
        Critical = 3
    End Enum

    Public Enum DeliverableTypes As Integer
        Goods = 0
        Services = 1
        Process = 2
        Plan = 10
        Budget = 11
        Report = 12
        Audit = 13
    End Enum
#End Region

#Region "Properties"
    Public Property DependencyType() As Integer
        Get
            Return intDependencyType
        End Get
        Set(ByVal value As Integer)
            intDependencyType = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property DependencyTypeText() As String
        Get
            Return LIST_DependencyTypes(Me.DependencyType)
        End Get
    End Property

    Public Property InputType() As Integer
        Get
            Return intInputType
        End Get
        Set(ByVal value As Integer)
            intInputType = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property InputTypeText() As String
        Get
            Return LIST_InputTypes(Me.InputType)
        End Get
    End Property

    Public Property ImportanceLevel() As Integer
        Get
            Return intImportanceLevel
        End Get
        Set(ByVal value As Integer)
            intImportanceLevel = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ImportanceLevelText() As String
        Get
            Return LIST_ImportanceLevels(Me.ImportanceLevel)
        End Get
    End Property

    Public Property DeliverableType() As Integer
        Get
            Return intDeliverableType
        End Get
        Set(ByVal value As Integer)
            intDeliverableType = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property DeliverableTypeText() As String
        Get
            Dim strDeliverableType As String = String.Empty
            strDeliverableType = LIST_DeliverableTypes.Find(AddressOf FindDeliverableTypeText).Value
            Return strDeliverableType
        End Get
    End Property

    Private Function FindDeliverableTypeText(ByVal selItem As IdValuePair) As Boolean
        If selItem.Id = intDeliverableType Then Return True Else Return False
    End Function

    Public Property Deliverables As String
        Get
            Return strDeliverables
        End Get
        Set(value As String)
            strDeliverables = value
        End Set
    End Property

    Public Property Supplier() As String
        Get
            Return strSupplier
        End Get
        Set(ByVal value As String)
            strSupplier = value
        End Set
    End Property

    Public Property DateExpected As Date
        Get
            Return datDateExpected
        End Get
        Set(value As Date)
            datDateExpected = value
        End Set
    End Property

    Public Property DateDelivered As Date
        Get
            Return datDateDelivered
        End Get
        Set(value As Date)
            datDateDelivered = value
        End Set
    End Property


#End Region
End Class
