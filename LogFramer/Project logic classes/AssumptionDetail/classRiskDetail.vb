Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class RiskDetail

    Private intRiskCategory As Integer
    Private intLikelihood As Integer
    Private intRiskImpact As Integer
    Private intRiskResponse As Integer

    Public Sub New()

    End Sub

#Region "Enumerations"
    Public Enum RiskCategories As Integer
        NotDefined = 0
        Operational = 1
        Financial = 2
        Objectives = 3
        Reputation = 4
        Other = 5
    End Enum

    Public Enum Likelihoods As Integer
        NotDefined = 0
        VeryUnlikely = 1
        Unlikely = 2
        Likely = 3
        VeryLikely = 4
    End Enum

    Public Enum RiskImpacts As Integer
        Unknown = 0
        VeryLow = 1
        Low = 2
        High = 3
        VeryHigh = 4
    End Enum

    Public Enum RiskResponses As Integer
        NotDefined = 0
        Avoiding = 1
        Reducing = 2
        Sharing = 3
        Transferring = 4
        Accepting = 5
    End Enum
#End Region

#Region "Properties"
    Public Property RiskCategory() As Integer
        Get
            Return intRiskCategory
        End Get
        Set(ByVal value As Integer)
            intRiskCategory = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RiskCategoryText() As String
        Get
            Return LIST_RiskCategories(Me.RiskCategory)
        End Get
    End Property

    Public Property Likelihood() As Integer
        Get
            Return intLikelihood
        End Get
        Set(ByVal value As Integer)
            intLikelihood = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property LikelihoodText() As String
        Get
            Return LIST_Likelihoods(Me.Likelihood)
        End Get
    End Property

    Public Property RiskImpact() As Integer
        Get
            Return intRiskImpact
        End Get
        Set(ByVal value As Integer)
            intRiskImpact = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RiskImpactText() As String
        Get
            Return LIST_RiskImpacts(Me.RiskImpact)
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RiskLevel() As Single
        Get
            Dim sngLevel As Single = Me.Likelihood * Me.RiskImpact / 16
            Return sngLevel
        End Get
    End Property

    Public Property RiskResponse() As Integer
        Get
            Return intRiskResponse
        End Get
        Set(ByVal value As Integer)
            intRiskResponse = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property RiskResponseText() As String
        Get
            Return LIST_RiskResponses(Me.RiskResponse)
        End Get
    End Property
#End Region
End Class
