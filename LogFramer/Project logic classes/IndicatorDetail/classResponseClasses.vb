Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ResponseClass
    Private intIdResponseClass As Integer = -1
    Private intIdIndicator As Integer
    Private strClassName As String
    Private dblValue As Single
    Private objGuid, objParentIndicatorGuid As Guid


    Public Sub New()

    End Sub

    Public Sub New(ByVal className As String)
        Me.ClassName = className
    End Sub

    Public Sub New(ByVal className As String, ByVal value As Double)
        Me.ClassName = className
        Me.Value = value
    End Sub

    Public Property idResponseClass As Integer
        Get
            Return intIdResponseClass
        End Get
        Set(ByVal value As Integer)
            intIdResponseClass = value
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

    Public Property Guid() As Guid
        Get
            If objGuid = Nothing Or objGuid = Guid.Empty Then objGuid = Guid.NewGuid
            Return objGuid
        End Get
        Set(ByVal value As Guid)
            objGuid = value
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

    Public Property ClassName() As String
        Get
            Return strClassName
        End Get
        Set(ByVal value As String)
            strClassName = value
        End Set
    End Property

    Public Property Value() As Double
        Get
            Return dblValue
        End Get
        Set(ByVal value As Double)
            dblValue = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemName() As String
        Get
            Return LANG_ResponseClass
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public Shared ReadOnly Property ItemNamePlural() As String
        Get
            Return LANG_ResponseClasses
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return ClassName
    End Function
End Class

Public Class ResponseClasses
    Inherits System.Collections.CollectionBase

    Public Event ResponseClassAdded(ByVal sender As Object, ByVal e As ResponseClassAddedEventArgs)

    Public ReadOnly Property ClassValuesSet() As Boolean
        Get
            Dim boolSet As Boolean
            For Each selClass As ResponseClass In Me.List
                If selClass.Value <> 0 Then
                    boolSet = True
                    Exit For
                End If
            Next
            Return boolSet
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub Add(ByVal responseclass As ResponseClass)
        List.Add(responseclass)
        RaiseEvent ResponseClassAdded(Me, New ResponseClassAddedEventArgs(responseclass))
    End Sub

    Public Sub Insert(ByVal intIndex As Integer, ByVal responseClass As ResponseClass)
        List.Insert(intIndex, responseClass)
        RaiseEvent ResponseClassAdded(Me, New ResponseClassAddedEventArgs(responseClass))
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal responseclass As ResponseClass)
        If Me.List.Contains(responseclass) Then
            Me.List.Remove(responseclass)
        End If
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As ResponseClass
        Get
            If index <= Me.Count - 1 Then
                Return CType(List.Item(index), ResponseClass)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal responseclass As ResponseClass) As Integer
        Return List.IndexOf(responseclass)
    End Function

    Public Function Contains(ByVal responseclass As ResponseClass) As Boolean
        Return List.Contains(responseclass)
    End Function

    Public Function GetMaximumScore(ByVal boolSum As Boolean) As Double
        Dim dblMaxScore As Double

        'If Me.List.Count > 0 Then dblMaxScore = CType(Me.List(0), ResponseClass).Value

        For Each selResponseClass As ResponseClass In Me.List
            If boolSum = False Then
                If selResponseClass.Value > dblMaxScore Then dblMaxScore = selResponseClass.Value
            Else
                dblMaxScore += selResponseClass.Value
            End If

        Next

        Return dblMaxScore
    End Function

    Public Function GetResponseClassByGuid(ByVal objGuid As Guid) As ResponseClass
        Dim selResponseClass As ResponseClass = Nothing
        For Each objResponseClass As ResponseClass In Me.List
            If objResponseClass.Guid = objGuid Then
                selResponseClass = objResponseClass
                Exit For
            End If
        Next
        Return selResponseClass
    End Function
End Class

Public Class ResponseClassAddedEventArgs
    Inherits EventArgs

    Public Property ResponseClass As ResponseClass

    Public Sub New(ByVal objResponseClass As ResponseClass)
        MyBase.New()

        Me.ResponseClass = objResponseClass
    End Sub
End Class

