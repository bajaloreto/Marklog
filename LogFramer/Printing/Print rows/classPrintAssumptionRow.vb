Public Class PrintAssumptionRow
    Private intSection As Integer
    Private intRowHeight As Integer
    Private boolClippedRow As Boolean
    Private strStructSortNumber As String
    Private objStruct As Struct
    Private intStructHeight As Integer
    Private intStructIndent As Integer
    Private strAssumptionSortNumber As String
    Private objAssumption As Assumption
    Private intAssumptionHeight As Integer
    Private intAssumptionIndent As Integer
    Private intReasonHeight As Integer
    Private intHowToValidateHeight As Integer
    Private intImpactHeight As Integer
    Private intResponseStrategyHeight As Integer
    Private intOwnerHeight As Integer
    Private bmReasonImage As System.Drawing.Bitmap
    Private bmHowToValidateImage As System.Drawing.Bitmap
    Private bmImpactImage As System.Drawing.Bitmap
    Private bmResponseStrategyImage As System.Drawing.Bitmap
    Private bmOwnerImage As System.Drawing.Bitmap
    Private strSectionTitle As String = String.Empty

    Public Sub New()

    End Sub

    Public Sub New(ByVal section As Integer, ByVal structsortnumber As String, ByVal struct As Struct)
        Me.Section = section
        Me.StructSortNumber = structsortnumber
        Me.Struct = struct
    End Sub

    Public Sub New(ByVal section As Integer, ByVal assumptionsortnumber As String, ByVal assumption As Assumption)
        Me.Section = section
        Me.AssumptionSortNumber = assumptionsortnumber
        Me.Assumption = assumption
    End Sub

    Public Property Section() As Integer
        Get
            Return intSection
        End Get
        Set(ByVal value As Integer)
            intSection = value
        End Set
    End Property

    Public Property RowHeight As Integer
        Get
            Return intRowHeight
        End Get
        Set(ByVal value As Integer)
            intRowHeight = value
        End Set
    End Property

    Public Property ClippedRow As Boolean
        Get
            Return boolClippedRow
        End Get
        Set(ByVal value As Boolean)
            boolClippedRow = value
        End Set
    End Property

    Public Property StructSortNumber() As String
        Get
            Return strStructSortNumber
        End Get
        Set(ByVal value As String)
            strStructSortNumber = value
        End Set
    End Property

    Public Property Struct() As Struct
        Get
            Return objStruct
        End Get
        Set(ByVal value As Struct)
            objStruct = value
        End Set
    End Property

    Public Property AssumptionSortNumber() As String
        Get
            Return strAssumptionSortNumber
        End Get
        Set(ByVal value As String)
            strAssumptionSortNumber = value
        End Set
    End Property

    Public Property Assumption() As Assumption
        Get
            Return objAssumption
        End Get
        Set(ByVal value As Assumption)
            objAssumption = value
        End Set
    End Property

    Public Property Reason() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.AssumptionDetail IsNot Nothing Then
                Return Me.Assumption.AssumptionDetail.Reason
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.AssumptionDetail IsNot Nothing Then
                Assumption.AssumptionDetail.Reason = value
            End If
        End Set
    End Property

    Public Property ReasonImage() As Bitmap
        Get
            Return bmReasonImage
        End Get
        Set(ByVal value As Bitmap)
            bmReasonImage = value
        End Set
    End Property

    Public Property ReasonHeight As Integer
        Get
            Return intReasonHeight
        End Get
        Set(ByVal value As Integer)
            intReasonHeight = value
        End Set
    End Property

    Public Property HowToValidate() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.AssumptionDetail IsNot Nothing Then
                Return Me.Assumption.AssumptionDetail.HowToValidate
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.AssumptionDetail IsNot Nothing Then
                Assumption.AssumptionDetail.HowToValidate = value
            End If
        End Set
    End Property

    Public Property HowToValidateImage() As Bitmap
        Get
            Return bmHowToValidateImage
        End Get
        Set(ByVal value As Bitmap)
            bmHowToValidateImage = value
        End Set
    End Property

    Public Property HowToValidateHeight As Integer
        Get
            Return intHowToValidateHeight
        End Get
        Set(ByVal value As Integer)
            intHowToValidateHeight = value
        End Set
    End Property

    Public ReadOnly Property Validated() As Boolean
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.AssumptionDetail IsNot Nothing Then
                Return Me.Assumption.AssumptionDetail.Validated
            Else
                Return False
            End If
        End Get
    End Property

    Public Property Impact() As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.Impact
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Assumption IsNot Nothing Then
                Me.Assumption.Impact = value
            End If
        End Set
    End Property

    Public Property ImpactImage() As Bitmap
        Get
            Return bmImpactImage
        End Get
        Set(ByVal value As Bitmap)
            bmImpactImage = value
        End Set
    End Property

    Public Property ImpactHeight As Integer
        Get
            Return intImpactHeight
        End Get
        Set(ByVal value As Integer)
            intImpactHeight = value
        End Set
    End Property

    Public Property ResponseStrategy() As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.ResponseStrategy
            Else
                Return String.Empty
            End If
        End Get
        Set(ByVal value As String)
            If Me.Assumption IsNot Nothing Then
                Assumption.ResponseStrategy = value
            End If
        End Set
    End Property

    Public Property ResponseStrategyImage() As Bitmap
        Get
            Return bmResponseStrategyImage
        End Get
        Set(ByVal value As Bitmap)
            bmResponseStrategyImage = value
        End Set
    End Property

    Public Property ResponseStrategyHeight As Integer
        Get
            Return intResponseStrategyHeight
        End Get
        Set(ByVal value As Integer)
            intResponseStrategyHeight = value
        End Set
    End Property

    Public ReadOnly Property Owner() As String
        Get
            If Me.Assumption IsNot Nothing Then
                Return Me.Assumption.Owner
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property OwnerImage() As Bitmap
        Get
            Return bmOwnerImage
        End Get
        Set(ByVal value As Bitmap)
            bmOwnerImage = value
        End Set
    End Property

    Public Property OwnerHeight As Integer
        Get
            Return intOwnerHeight
        End Get
        Set(ByVal value As Integer)
            intOwnerHeight = value
        End Set
    End Property
End Class

Public Class PrintAssumptionRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintAssumptionRow)
        List.Add(row)
    End Sub

    Public Sub AddRange(ByVal objPrintAssumptionRows As PrintAssumptionRows)
        InnerList.AddRange(objPrintAssumptionRows)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintAssumptionRow)
        If index > Count Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        ElseIf index = Count Then
            List.Add(row)
        Else
            List.Insert(index, row)
        End If
    End Sub

    Public Sub Remove(ByVal index As Integer)
        If index > Count - 1 Or index < 0 Then
            System.Windows.Forms.MessageBox.Show(ERR_IndexNotValid)
        Else
            List.RemoveAt(index)
        End If
    End Sub

    Public Sub Remove(ByVal row As PrintAssumptionRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("RiskTableRow not in list!")
        Else
            List.Remove(row)
        End If
    End Sub

    Public Sub RemoveRange(ByVal index As Integer, ByVal count As Integer)
        For i = index To index + count - 1
            If i <= List.Count - 1 Then
                List.RemoveAt(i)
            End If
        Next
    End Sub

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintAssumptionRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintAssumptionRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintAssumptionRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintAssumptionRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
