Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Statement
    Inherits LogframeObject

    Private intIdStatement As Integer = -1
    Private intIdIndicator As Integer
    Private strFirstLabel As String
    Private strSecondLabel As String
    Private objGuid, objParentIndicatorGuid As Guid

    <ScriptIgnore()> _
    Public WithEvents ValuesDetail As New ValuesDetail

#Region "Properties"
    Public Property idStatement As Integer
        Get
            Return intIdStatement
        End Get
        Set(ByVal value As Integer)
            intIdStatement = value
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

    Public Property FirstLabel() As String
        Get
            Return strFirstLabel
        End Get
        Set(ByVal value As String)
            strFirstLabel = value
        End Set
    End Property

    Public Property SecondLabel() As String
        Get
            Return strSecondLabel
        End Get
        Set(ByVal value As String)
            strSecondLabel = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_Statement
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_Statements
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub New()

    End Sub

    Public Sub New(ByVal firstlabel As String)
        Me.FirstLabel = firstlabel
    End Sub

    Public Sub New(ByVal firstlabel As String, ByVal unit As String, ByVal nrdecimals As Integer, Optional ByVal valuename As String = "")
        Me.FirstLabel = firstlabel

        With ValuesDetail
            .Unit = unit
            .NrDecimals = nrdecimals
            .ValueName = valuename
        End With
    End Sub

    Protected Overrides Sub OnGuidChanged()
        'no child collections
    End Sub

    Public Overrides Function ToString() As String
        If String.IsNullOrEmpty(Me.FirstLabel) = False Then
            Return Me.FirstLabel
        Else
            Return String.Empty
        End If

    End Function
#End Region

    
End Class

Public Class Statements
    Inherits System.Collections.CollectionBase

    Public Event StatementAdded(ByVal sender As Object, ByVal e As StatementAddedEventArgs)

    Public Sub Add(ByVal statement As Statement)
        If statement IsNot Nothing Then
            List.Add(statement)
            RaiseEvent StatementAdded(Me, New StatementAddedEventArgs(statement))
        End If
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal statement As Statement)
        If statement IsNot Nothing Then
            List.Insert(intIndex, statement)
            RaiseEvent StatementAdded(Me, New StatementAddedEventArgs(statement))
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal statement As Statement)
        If Me.List.Contains(statement) Then
            Me.List.Remove(statement)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As Statement
        Get
            If index >= 0 And index <= Me.List.Count - 1 Then
                Return CType(List.Item(index), Statement)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal statement As Statement) As Integer
        Return List.IndexOf(statement)
    End Function

    Public Function Contains(ByVal statement As Statement) As Boolean
        Return List.Contains(statement)
    End Function

    Public Function GetStatementByGuid(ByVal objGuid As Guid) As Statement
        Dim selStatement As Statement = Nothing
        For Each objStatement As Statement In Me.List
            If objStatement.Guid = objGuid Then
                selStatement = objStatement
                Exit For
            End If
        Next
        Return selStatement
    End Function
End Class

Public Class StatementAddedEventArgs
    Inherits EventArgs

    Public Property Statement As Statement

    Public Sub New(ByVal objStatement As Statement)
        MyBase.New()

        Me.Statement = objStatement
    End Sub
End Class