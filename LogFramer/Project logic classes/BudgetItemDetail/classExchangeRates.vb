Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ExchangeRate
    Private intIdExchangeRate As Integer = -1
    Private intIdBudget As Integer
    Private strCurrencyCode As String
    Private sngExchangeRate As Single = 1
    Private objGuid, objParentBudgetGuid As Guid


    Public Sub New()

    End Sub

    Public Sub New(ByVal currencycode As String, ByVal exchangerate As Single)
        Me.CurrencyCode = currencycode
        Me.ExchangeRate = exchangerate
    End Sub

    Public Property idExchangeRate As Integer
        Get
            Return intIdExchangeRate
        End Get
        Set(ByVal value As Integer)
            intIdExchangeRate = value
        End Set
    End Property

    Public Property idBudget As Integer
        Get
            Return intIdBudget
        End Get
        Set(ByVal value As Integer)
            intIdBudget = value
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

    Public Property ParentBudgetGuid() As Guid
        Get
            Return objParentBudgetGuid
        End Get
        Set(ByVal value As Guid)
            objParentBudgetGuid = value
        End Set
    End Property

    Public Property CurrencyCode() As String
        Get
            Return strCurrencyCode
        End Get
        Set(ByVal value As String)
            strCurrencyCode = value
        End Set
    End Property

    Public Property ExchangeRate() As Single
        Get
            Return sngExchangeRate
        End Get
        Set(ByVal value As Single)
            sngExchangeRate = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_ExchangeRate
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_ExchangeRates
        End Get
    End Property
End Class

Public Class ExchangeRates
    Inherits System.Collections.CollectionBase

    Public Event ExchangeRateAdded(ByVal sender As Object, ByVal e As ExchangeRateAddedEventArgs)

    Public Sub New()

    End Sub

    Public Sub Add(ByVal exchangerate As ExchangeRate)
        List.Add(exchangerate)
        RaiseEvent ExchangeRateAdded(Me, New ExchangeRateAddedEventArgs(exchangerate))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal responseClass As ExchangeRate)
        List.Insert(intIndex, responseClass)
        RaiseEvent ExchangeRateAdded(Me, New ExchangeRateAddedEventArgs(responseClass))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal exchangerate As ExchangeRate)
        If Me.List.Contains(exchangerate) Then
            Me.List.Remove(exchangerate)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ExchangeRate
        Get
            If index <= Me.Count - 1 Then
                Return CType(List.Item(index), ExchangeRate)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal exchangerate As ExchangeRate) As Integer
        Return List.IndexOf(exchangerate)
    End Function

    Public Function Contains(ByVal exchangerate As ExchangeRate) As Boolean
        Return List.Contains(exchangerate)
    End Function

    Public Function GetExchangeRate(ByVal strCurrencyCode As String) As Double
        Dim dblExchangeRate As Double = 1
        For Each objExchangeRate As ExchangeRate In Me.List
            If objExchangeRate.CurrencyCode = strCurrencyCode Then
                dblExchangeRate = objExchangeRate.ExchangeRate
                Exit For
            End If
        Next
        Return dblExchangeRate
    End Function

    Public Function GetExchangeRateByGuid(ByVal objGuid As Guid) As ExchangeRate
        Dim selExchangeRate As ExchangeRate = Nothing
        For Each objExchangeRate As ExchangeRate In Me.List
            If objExchangeRate.Guid = objGuid Then
                selExchangeRate = objExchangeRate
                Exit For
            End If
        Next
        Return selExchangeRate
    End Function
End Class

Public Class ExchangeRateAddedEventArgs
    Inherits EventArgs

    Public Property ExchangeRate As ExchangeRate

    Public Sub New(ByVal objExchangeRate As ExchangeRate)
        MyBase.New()

        Me.ExchangeRate = objExchangeRate
    End Sub
End Class

