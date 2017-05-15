Imports System.Xml.Serialization
Imports System.Web.Script.Serialization

Public Class ValuesDetail
    Private strUnit As String
    Private intNrDecimals As Integer
    Private strValueName As String
    Private strFormula As String

    <ScriptIgnore()> _
    Public WithEvents ValueRange As New ValueRange

    Public Sub New()

    End Sub

    Public Sub New(ByVal valuename As String, ByVal nrdecimals As Integer, ByVal unit As String)
        Me.ValueName = valuename
        Me.NrDecimals = nrdecimals
        Me.Unit = unit
    End Sub

#Region "Properties"
    Public Property NrDecimals() As Integer
        Get
            Return intNrDecimals
        End Get
        Set(ByVal value As Integer)
            intNrDecimals = value
        End Set
    End Property

    Public Property Unit() As String
        Get
            Return strUnit
        End Get
        Set(ByVal value As String)
            strUnit = value
        End Set
    End Property

    Public Property ValueName() As String
        Get
            If String.IsNullOrEmpty(strValueName) Then
                Return "x"
            Else
                Return strValueName
            End If
        End Get
        Set(ByVal value As String)
            strValueName = value
        End Set
    End Property

    Public Property Formula As String
        Get
            Return strFormula
        End Get
        Set(value As String)
            strFormula = value
        End Set
    End Property

    Public ReadOnly Property FormatString As String
        Get
            Dim strPrecision As String = "#,##0."
            If Me.NrDecimals > 0 Then
                For i = 1 To Me.NrDecimals
                    strPrecision &= "0"
                Next
            End If
            If String.IsNullOrEmpty(Me.Unit) = False Then strPrecision = String.Format("{0} '{1}'", strPrecision, Me.Unit)

            Return strPrecision
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function DisplayField(ByVal selRange As ValueRange) As String
        Dim strRange As String = String.Empty, strOp As String

        If String.IsNullOrEmpty(selRange.OpMin) = False And String.IsNullOrEmpty(selRange.OpMax) = False Then
            If selRange.OpMin = CONST_Equals Then strOp = CONST_Equals Else strOp = CONST_SmallerThan
            If Len(selRange.OpMin) > 1 Then strOp &= CONST_Equals

            strRange = String.Format("{0} {1}{2}{3} {4}", DisplayAsUnit(selRange.MinValue), strOp, Me.ValueName, selRange.OpMax, DisplayAsUnit(selRange.MaxValue))
        Else
            If String.IsNullOrEmpty(selRange.OpMin) = False Then
                strRange = Me.ValueName & " " & selRange.OpMin & DisplayAsUnit(selRange.MinValue)
                strRange = String.Format("{0} {1}{2}", Me.ValueName, selRange.OpMin, DisplayAsUnit(selRange.MinValue))
            ElseIf String.IsNullOrEmpty(selRange.OpMax) = False Then
                strRange = String.Format("{0} {1}{2}", Me.ValueName, selRange.OpMax, DisplayAsUnit(selRange.MaxValue))
            Else
                strRange = "..."
            End If
        End If

        Return strRange
    End Function

    Private Function DisplayAsUnit(ByVal sngValue As Single) As String
        Dim strPrecision As String = "#,##0."
        If Me.NrDecimals > 0 Then
            For i = 1 To Me.NrDecimals
                strPrecision &= "0"
            Next
        End If
        If String.IsNullOrEmpty(Me.Unit) = False Then strPrecision = String.Format("{0} '{1}'", strPrecision, Me.Unit)
        Return sngValue.ToString(strPrecision)
    End Function
#End Region
End Class
