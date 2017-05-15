Public Class PrintDependencyRow
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
    Private intDependencyTypeHeight As Integer
    Private intInputTypeHeight As Integer
    Private intImportanceLevelHeight As Integer
    Private intImpactHeight As Integer
    Private intSupplierHeight As Integer
    Private intDeliverableTypeHeight As Integer
    Private intDeliverablesHeight As Integer
    Private intResponseStrategyHeight As Integer
    Private intOwnerHeight As Integer
    Private bmDependencytypeImage As System.Drawing.Bitmap
    Private bmInputTypeImage As System.Drawing.Bitmap
    Private bmImportanceLevelImage As System.Drawing.Bitmap
    Private bmImpactImage As System.Drawing.Bitmap
    Private bmSupplierImage As System.Drawing.Bitmap
    Private bmDeliverableType As System.Drawing.Bitmap
    Private bmDeliverablesImage As System.Drawing.Bitmap
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

    Public ReadOnly Property DependencyTypeText() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.DependencyTypeText
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property DependencyTypeImage() As Bitmap
        Get
            Return bmDependencyTypeImage
        End Get
        Set(ByVal value As Bitmap)
            bmDependencyTypeImage = value
        End Set
    End Property

    Public Property DependencyTypeHeight As Integer
        Get
            Return intDependencyTypeHeight
        End Get
        Set(ByVal value As Integer)
            intDependencyTypeHeight = value
        End Set
    End Property

    Public ReadOnly Property InputTypeText() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.InputTypeText
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property InputTypeImage() As Bitmap
        Get
            Return bmInputTypeImage
        End Get
        Set(ByVal value As Bitmap)
            bmInputTypeImage = value
        End Set
    End Property

    Public Property InputTypeHeight As Integer
        Get
            Return intInputTypeHeight
        End Get
        Set(ByVal value As Integer)
            intInputTypeHeight = value
        End Set
    End Property

    Public ReadOnly Property ImportanceLevelText() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.ImportanceLevelText
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property ImportanceLevelImage() As Bitmap
        Get
            Return bmImportanceLevelImage
        End Get
        Set(ByVal value As Bitmap)
            bmImportanceLevelImage = value
        End Set
    End Property

    Public Property ImportanceLevelHeight As Integer
        Get
            Return intImportanceLevelHeight
        End Get
        Set(ByVal value As Integer)
            intImportanceLevelHeight = value
        End Set
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

    Public ReadOnly Property Supplier() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.Supplier
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property SupplierImage() As Bitmap
        Get
            Return bmSupplierImage
        End Get
        Set(ByVal value As Bitmap)
            bmSupplierImage = value
        End Set
    End Property

    Public Property SupplierHeight As Integer
        Get
            Return intSupplierHeight
        End Get
        Set(ByVal value As Integer)
            intSupplierHeight = value
        End Set
    End Property

    Public ReadOnly Property DeliverableTypeText() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.DeliverableTypeText
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property DeliverableTypeImage() As Bitmap
        Get
            Return bmDeliverableType
        End Get
        Set(ByVal value As Bitmap)
            bmDeliverableType = value
        End Set
    End Property

    Public Property DeliverableTypeHeight As Integer
        Get
            Return intDeliverableTypeHeight
        End Get
        Set(ByVal value As Integer)
            intDeliverableTypeHeight = value
        End Set
    End Property

    Public ReadOnly Property DateExpected() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing And _
                Me.Assumption.DependencyDetail.DateExpected <> Date.MinValue Then

                Return Me.Assumption.DependencyDetail.DateExpected.ToString("d")
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property DateDelivered() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing And _
                Me.Assumption.DependencyDetail.DateDelivered <> Date.MinValue Then

                Return Me.Assumption.DependencyDetail.DateDelivered.ToString("d")
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public ReadOnly Property Deliverables() As String
        Get
            If Me.Assumption IsNot Nothing AndAlso Me.Assumption.DependencyDetail IsNot Nothing Then
                Return Me.Assumption.DependencyDetail.Deliverables
            Else
                Return String.Empty
            End If
        End Get
    End Property

    Public Property DeliverablesImage() As Bitmap
        Get
            Return bmDeliverablesImage
        End Get
        Set(ByVal value As Bitmap)
            bmDeliverablesImage = value
        End Set
    End Property

    Public Property DeliverablesHeight As Integer
        Get
            Return intDeliverablesHeight
        End Get
        Set(ByVal value As Integer)
            intDeliverablesHeight = value
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

Public Class PrintDependencyRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintDependencyRow)
        List.Add(row)
    End Sub

    Public Sub AddRange(ByVal objPrintDependencyRows As PrintDependencyRows)
        InnerList.AddRange(objPrintDependencyRows)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintDependencyRow)
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

    Public Sub Remove(ByVal row As PrintDependencyRow)
        If List.Contains(row) = False Then
            System.Windows.Forms.MessageBox.Show("Dependency Row not in list!")
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

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintDependencyRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintDependencyRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintDependencyRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintDependencyRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
