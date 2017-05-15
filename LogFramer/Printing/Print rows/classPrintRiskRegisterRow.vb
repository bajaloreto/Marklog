Public Class PrintRiskRegisterRow
    Private intRowHeight As Integer
    Private boolClippedRow As Boolean
    Private strAssumptionSortNumber As String
    Private objAssumption As Assumption
    Private intAssumptionHeight As Integer
    Private intAssumptionIndent As Integer
    Private strRiskResponse As String = String.Empty
    Private strRiskResponseText As String = String.Empty
    Private intRiskResponseHeight As Integer
    Private bmRiskResponseImage As System.Drawing.Bitmap
    Private strStructSortNumber As String
    Private objStruct As Struct
    Private intStructHeight As Integer
    Private intStructIndent As Integer
    Private strRiskLevel As String = String.Empty
    Private intRiskLevelHeight As Integer
    Private bmRiskLevelImage As System.Drawing.Bitmap
   
    Private strLikelihoodText As String = String.Empty
    Private strRiskImpactText As String = String.Empty
    Private sngRiskLevel As Single
    Private strSectionTitle As String = String.Empty

    Public Sub New()

    End Sub

    Public Sub New(ByVal sectiontitle As String)
        Me.SectionTitle = sectiontitle
    End Sub

    Public Sub New(ByVal assumptionsortnumber As String, ByVal assumption As Assumption, ByVal structsortnumber As String, ByVal struct As Struct)
        Me.AssumptionSortNumber = assumptionsortnumber
        Me.Assumption = assumption
        Me.StructSortNumber = structsortnumber
        Me.Struct = struct

        SetRiskResponseText()
        SetRiskLevelText()
    End Sub

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

    Public Property RiskResponse() As String
        Get
            Return strRiskResponse
        End Get
        Set(ByVal value As String)
            strRiskResponse = value
        End Set
    End Property

    Public Property RiskResponseImage() As Bitmap
        Get
            Return bmRiskResponseImage
        End Get
        Set(ByVal value As Bitmap)
            bmRiskResponseImage = value
        End Set
    End Property

    Public Property RiskResponseHeight As Integer
        Get
            Return intRiskResponseHeight
        End Get
        Set(ByVal value As Integer)
            intRiskResponseHeight = value
        End Set
    End Property

    Public Sub SetRiskResponseText()
        If objAssumption IsNot Nothing AndAlso objAssumption.RiskDetail IsNot Nothing Then
            Dim strStatement As String
            strStatement = objAssumption.RiskDetail.RiskResponseText

            If String.IsNullOrEmpty(strStatement) = False Then strStatement &= vbLf
            strStatement &= objAssumption.ResponseStrategy

            Me.RiskResponse = TextToRichText(strStatement)
        End If
    End Sub

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

    Public Property RiskLevel() As String
        Get
            Return strRiskLevel
        End Get
        Set(ByVal value As String)
            strRiskLevel = value
        End Set
    End Property

    Public Property RiskLevelImage() As Bitmap
        Get
            Return bmRiskLevelImage
        End Get
        Set(ByVal value As Bitmap)
            bmRiskLevelImage = value
        End Set
    End Property

    Public Property RiskLevelHeight As Integer
        Get
            Return intRiskLevelHeight
        End Get
        Set(ByVal value As Integer)
            intRiskLevelHeight = value
        End Set
    End Property

    Public Sub SetRiskLevelText()
        If objAssumption IsNot Nothing AndAlso objAssumption.RiskDetail IsNot Nothing Then
            Dim strText As String = String.Empty

            If String.IsNullOrEmpty(objAssumption.RiskDetail.RiskImpactText) = False Then
                With objAssumption.RiskDetail
                    Dim strImpactText As String = .RiskImpactText
                    Dim strSplit() As String = strImpactText.Split("-"c)
                    strImpactText = Trim(strSplit(0))

                    strText = String.Format("{1} {2}{0}{3} {4}{0}{5} {6}", vbCrLf, _
                                            ToLabel(LANG_Likelihood), .LikelihoodText, _
                                            ToLabel(LANG_Impact), strImpactText, _
                                            ToLabel(LANG_RiskLevel), .RiskLevel.ToString("P0"))
                    Me.RiskLevel = TextToRichText(strText)
                End With
            End If

        End If
    End Sub

    Public Property SectionTitle As String
        Get
            Return strSectionTitle
        End Get
        Set(ByVal value As String)
            strSectionTitle = value
        End Set
    End Property
End Class

Public Class PrintRiskRegisterRows
    Inherits System.Collections.CollectionBase

    Public Sub New()

    End Sub

    Public Sub Add(ByVal row As PrintRiskRegisterRow)
        List.Add(row)
    End Sub

    Public Sub AddRange(ByVal objPrintRiskRegisterRows As PrintRiskRegisterRows)
        InnerList.AddRange(objPrintRiskRegisterRows)
    End Sub

    Public Sub Insert(ByVal index As Integer, ByVal row As PrintRiskRegisterRow)
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

    Public Sub Remove(ByVal row As PrintRiskRegisterRow)
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

    Default Public ReadOnly Property Item(ByVal index As Integer) As PrintRiskRegisterRow
        Get
            If index > Count - 1 Or index < 0 Then
                Return Nothing
            Else
                Return CType(List.Item(index), PrintRiskRegisterRow)
            End If
        End Get
    End Property

    Public Function IndexOf(ByVal row As PrintRiskRegisterRow) As Integer
        Return List.IndexOf(row)
    End Function

    Public Function Contains(ByVal row As PrintRiskRegisterRow) As Boolean
        Return List.Contains(row)
    End Function
End Class
