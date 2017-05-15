Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class Struct
    Inherits LogframeObject

    Private intIdStruct As Integer = -1
    Private intIdParent As Integer

    <ScriptIgnore()> _
    Public WithEvents Indicators As New Indicators

    <ScriptIgnore()> _
    Public WithEvents Assumptions As New Assumptions

#Region "Properties"
    Public Overridable Property idStruct As Integer
        Get
            Return intIdStruct
        End Get
        Set(ByVal value As Integer)
            intIdStruct = value
        End Set
    End Property

    Public Property idParent As Integer
        Get
            Return intIdParent
        End Get
        Set(ByVal value As Integer)
            intIdParent = value
        End Set
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property ResourceBudget() As Boolean
        Get
            If Me.Section = 4 And CurrentProjectForm.dgvLogframe.ShowResourcesBudget = True Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    <XmlIgnore()> _
    <ScriptIgnore()> _
    Public ReadOnly Property HasHorizontalChildren() As Boolean
        Get
            If Me.ResourceBudget = False Then
                If Me.Indicators.Count = 0 And Me.Assumptions.Count = 0 Then Return False Else Return True
            Else
                Dim selActivity As Activity = TryCast(Me, Activity)

                If selActivity IsNot Nothing Then
                    If selActivity.Resources.Count = 0 And Me.Assumptions.Count = 0 Then Return False Else Return True
                End If
            End If
        End Get
    End Property
#End Region

#Region "Methods & Events"
    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub

    Public Function GetItemName(Optional ByVal section As Integer = 0) As String
        If section > 0 Then Me.Section = section

        Select Case Me.Section
            Case SectionTypes.Goal
                Return My.Settings.setStruct1sing
            Case SectionTypes.Purpose
                Return My.Settings.setStruct2sing
            Case SectionTypes.Output
                Return My.Settings.setStruct3sing
            Case SectionTypes.Activity
                Return My.Settings.setStruct4sing
        End Select
        Return Nothing
    End Function

    Public Function GetItemNamePlural(Optional ByVal section As Integer = 0) As String
        If section > 0 Then Me.Section = section

        Select Case Me.Section
            Case SectionTypes.Goal
                Return My.Settings.setStruct1
            Case SectionTypes.Purpose
                Return My.Settings.setStruct2
            Case SectionTypes.Output
                Return My.Settings.setStruct3
            Case SectionTypes.Activity
                Return My.Settings.setStruct4
        End Select
        Return Nothing
    End Function

    Private Sub Indicators_IndicatorAdded(ByVal sender As Object, ByVal e As IndicatorAddedEventArgs) Handles Indicators.IndicatorAdded
        Dim selIndicator As Indicator = e.Indicator

        selIndicator.idParent = Me.idStruct
        selIndicator.ParentStructGuid = Me.Guid
        selIndicator.Section = Me.Section
    End Sub

    Private Sub Assumptions_AssumptionAdded(ByVal sender As Object, ByVal e As AssumptionAddedEventArgs) Handles Assumptions.AssumptionAdded
        Dim selAssumption As Assumption = e.Assumption

        selAssumption.idStruct = Me.idStruct
        selAssumption.ParentStructGuid = Me.Guid
    End Sub

    Protected Overrides Sub OnGuidChanged() Handles Me.GuidChanged
        For Each selIndicator As Indicator In Me.Indicators
            selIndicator.ParentStructGuid = Me.Guid
        Next
        For Each selAssumption As Assumption In Me.Assumptions
            selAssumption.ParentStructGuid = Me.Guid
        Next
    End Sub
#End Region
End Class

Public Class Structs
    Inherits System.Collections.CollectionBase

#Region "Properties"
    Default Public ReadOnly Property Item(ByVal index As Integer) As Struct
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), Struct)
            End If
        End Get
    End Property
#End Region

#Region "General methods"
    Public Sub New()

    End Sub

    Public Sub Add(ByVal struct As Struct)
        If struct IsNot Nothing Then
            List.Add(struct)
        End If
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal struct As Struct)
        If struct IsNot Nothing Then
            If index > Count Or index < 0 Then
                System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
            ElseIf index = Count Then
                List.Add(struct)
            Else
                List.Insert(index, struct)
            End If
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal struct As Struct)
        If List.Contains(struct) = False Then
            System.Windows.Forms.MessageBox.Show("Object not in list!")
        Else
            List.Remove(struct)
        End If
    End Sub

    Public Function IndexOf(ByVal struct As Struct) As Integer
        Return List.IndexOf(struct)
    End Function

    Public Function Contains(ByVal struct As Struct) As Boolean
        Return List.Contains(struct)
    End Function
#End Region

#Region "Get by GUID"
    Public Function GetIndicatorByGuid(ByVal objGuid As Guid) As Indicator
        Dim selIndicator As Indicator = Nothing
        For Each selStruct As Struct In Me.List
            selIndicator = selStruct.Indicators.GetIndicatorByGuid(objGuid)
            If selIndicator IsNot Nothing Then Exit For
        Next
        Return selIndicator
    End Function

    Public Function GetStatementByGuid(ByVal objGuid As Guid) As Statement
        Dim selStatement As Statement = Nothing
        For Each selStruct As Struct In Me.List
            selStatement = selStruct.Indicators.GetStatementByGuid(objGuid)

            If selStatement IsNot Nothing Then Return selStatement
        Next
        Return selStatement
    End Function

    Public Function GetTargetByGuid(ByVal objGuid As Guid) As Target
        Dim selTarget As Target = Nothing
        For Each selStruct As Struct In Me.List
            selTarget = selStruct.Indicators.GetTargetByGuid(objGuid)

            If selTarget IsNot Nothing Then Return selTarget
        Next
        Return selTarget
    End Function

    'Public Function GetTargetBooleanValueByGuid(ByVal objGuid As Guid) As BooleanValue
    '    Dim selTargetBooleanValue As BooleanValue = Nothing
    '    For Each selStruct As Struct In Me.List
    '        selTargetBooleanValue = selStruct.Indicators.GetTargetBooleanValueByGuid(objGuid)

    '        If selTargetBooleanValue IsNot Nothing Then Return selTargetBooleanValue
    '    Next
    '    Return selTargetBooleanValue
    'End Function

    'Public Function GetIntegerResponseByGuid(ByVal objGuid As Guid) As IntegerResponse
    '    Dim selIntegerResponse As IntegerResponse = Nothing
    '    For Each selStruct As Struct In Me.List
    '        selIntegerResponse = selStruct.Indicators.GetIntegerResponseByGuid(objGuid)

    '        If selIntegerResponse IsNot Nothing Then Return selIntegerResponse
    '    Next
    '    Return selIntegerResponse
    'End Function

    Public Function GetResponseClassByGuid(ByVal objGuid As Guid) As ResponseClass
        Dim selResponseClass As ResponseClass = Nothing
        For Each selStruct As Struct In Me.List
            selResponseClass = selStruct.Indicators.GetResponseClassByGuid(objGuid)

            If selResponseClass IsNot Nothing Then Return selResponseClass
        Next
        Return selResponseClass
    End Function

    Public Function GetVerificationSourceByGuid(ByVal objGuid As Guid) As VerificationSource
        Dim selVerificationSource As VerificationSource = Nothing
        For Each selStruct As Struct In Me.List
            selVerificationSource = selStruct.Indicators.GetVerificationSourceByGuid(objGuid)

            If selVerificationSource IsNot Nothing Then Return selVerificationSource
        Next
        Return selVerificationSource
    End Function

    Public Function GetAssumptionByGuid(ByVal objGuid As Guid) As Assumption
        Dim selAssumption As Assumption = Nothing
        For Each selStruct As Struct In Me.List
            selAssumption = selStruct.Assumptions.GetAssumptionByGuid(objGuid)
            If selAssumption IsNot Nothing Then Exit For
        Next
        Return selAssumption
    End Function
#End Region

End Class

Public Class SectionTitle
    Inherits Struct

    <ScriptIgnore()> _
    Public WithEvents Resources As New Resources

    Public Sub New()

    End Sub

    Public Sub New(ByVal RTF As String)
        Me.RTF = RTF
    End Sub
End Class

Public Class SectionTitles
    Inherits Structs

End Class
