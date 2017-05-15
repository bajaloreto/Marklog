Public Class DataGridViewProjectLogic
    Inherits DataGridView

    Public Schemes As New List(Of ProjectLogicScheme)
    Private objScheme As ProjectLogicScheme
    Private intSelectedSchemeIndex As Integer

    Public Property SelectedSchemeIndex As Integer
        Get
            If Me.SelectedColumns(0).Index >= 0 Then
                intSelectedSchemeIndex = Me.SelectedColumns(0).Index
            Else
                intSelectedSchemeIndex = 0
            End If

            Return intSelectedSchemeIndex
        End Get
        Set(value As Integer)
            If value >= 0 And value <= Me.ColumnCount - 1 Then
                intSelectedSchemeIndex = value
                Me.CurrentCell = Me(intSelectedSchemeIndex, 0)
            Else
                intSelectedSchemeIndex = 0
            End If
        End Set
    End Property

    Public ReadOnly Property SelectedScheme As ProjectLogicScheme
        Get
            If Me.SelectedColumns.Count > 0 Then
                objScheme = Me.Schemes(SelectedSchemeIndex)
            End If

            Return objScheme
        End Get
    End Property

    Public Sub New()
        Me.AllowUserToAddRows = False
        Me.AllowUserToDeleteRows = False
        Me.BackgroundColor = Color.White
        Me.ReadOnly = True
        Me.RowHeadersVisible = False
        Me.SelectionMode = DataGridViewSelectionMode.FullColumnSelect
        Me.MultiSelect = False

        With ColumnHeadersDefaultCellStyle
            .Alignment = DataGridViewContentAlignment.MiddleCenter
            .Font = New Font(.Font.FontFamily, FontStyle.Bold)
            .WrapMode = DataGridViewTriState.True
        End With
    End Sub

    Public Sub LoadColumns(ByVal selectedschemeindex As Integer)
        LoadSchemes()

        For i = 1 To Schemes.Count
            Dim colScheme As New DataGridViewTextBoxColumn
            With colScheme
                .Name = String.Format("{0} {1}", LANG_Scheme, i)
                .Width = 150
                .SortMode = DataGridViewColumnSortMode.NotSortable
            End With
            Me.Columns.Add(colScheme)
        Next
        Me.RowCount = 4
        With Me.Rows(0)
            For i = 0 To Schemes.Count - 1
                .Cells(i).Value = Schemes(i).GoalNamePlural
            Next
        End With
        With Me.Rows(1)
            For i = 0 To Schemes.Count - 1
                .Cells(i).Value = Schemes(i).PurposeNamePlural
            Next
        End With
        With Me.Rows(2)
            For i = 0 To Schemes.Count - 1
                .Cells(i).Value = Schemes(i).OutputNamePlural
            Next
        End With
        With Me.Rows(3)
            For i = 0 To Schemes.Count - 1
                .Cells(i).Value = Schemes(i).ActivityNamePlural
            Next
        End With

        Me.SelectedSchemeIndex = selectedschemeindex
        
    End Sub

    Private Sub LoadSchemes()
        Schemes.Clear()
        Select Case UserLanguage
            Case "fr"
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs généraux", "Objectif général", "Objectifs spécifiques", "Objectif spécifique", "Résultats", "Résultat", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs généraux", "Objectif général", "Objectifs spécifiques", "Objectif spécifique", "Résultats attendus", "Résultat attendu", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Impact", "Impact", "Effets", "Effet", "Extrants", "Extrant", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs finaux", "Objectif final", "Objectifs intermédiaires", "Objectif intermédiaire", "Extrants", "Extrant", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs stratégiques", "Objectif stratégique", "Résultats intermédiaires", "Résultat intermédiaire", "Extrants", "Extrant", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs de développement", "Objectif de développement", "Objectifs immédiats", "Objectif immédiat", "Extrants", "Extrant", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs à long terme", "Objectif à long terme", "Objectifs à court terme", "Objectif à court terme", "Extrants", "Extrant", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs finaux", "Objectif final", "Objectifs stratégiques", "Objectif stratégique", "Résultats intermédiaires", "Résultat intermédiaire", "Activités", "Activité"))
                Schemes.Add(New ProjectLogicScheme(False, "Objectifs globaux", "Objectif global", "Objectifs du projet", "Objectif du projet", "Résultats", "Résultat", "Activités", "Activité"))
            Case Else
                Schemes.Add(New ProjectLogicScheme(False, "Goals", "Goal", "Purposes", "Purpose", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Impact", "Impact", "Outcomes", "Outcome", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Impact", "Impact", "Effects", "Effect", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Final goals", "Final goal", "Intermediate objectives", "Intermediate objective", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Strategic objectives", "Strategic objective", "Intermediate results", "Intermediate result", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Development objectives", "Development objective", "Immediate objectives", "Immediate objective", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Long-term objectives", "Long-term objective", "Short-term objectives", "Short-term objective", "Outputs", "Output", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "General objectives", "General objective", "Specific objectives", "Specific objective", "Results", "Result", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Final goals", "Final goal", "Strategic objectives", "Strategic objective", "Intermediate results", "Intermediate result", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Overall goals", "Overall goal", "Project purposes", "Project purpose", "Results", "Result", "Activities", "Activity"))
                Schemes.Add(New ProjectLogicScheme(False, "Overall objectives", "Overall objective", "Project purposes", "Project purpose", "Results", "Result", "Activities", "Activity"))
        End Select

    End Sub
End Class

Public Class ProjectLogicScheme
    Private boolSelected As Boolean
    Private strGoalNamePlural As String
    Private strGoalName As String
    Private strPurposeNamePlural As String
    Private strPurposeName As String
    Private strOutputNamePlural As String
    Private strOutputName As String
    Private strActivityNamePlural As String
    Private strActivityName As String

    Public Sub New()

    End Sub

    Public Sub New(ByVal selected As Boolean, ByVal goalnameplural As String, ByVal goalname As String, _
                   ByVal purposenameplural As String, ByVal purposename As String, _
                   ByVal outputnameplural As String, ByVal outputname As String, _
                   ByVal activitynameplural As String, ByVal activityname As String)
        Me.Selected = selected
        Me.GoalNamePlural = goalnameplural
        Me.GoalName = goalname
        Me.PurposeNamePlural = purposenameplural
        Me.PurposeName = purposename
        Me.OutputNamePlural = outputnameplural
        Me.OutputName = outputname
        Me.ActivityNamePlural = activitynameplural
        Me.ActivityName = activityname
    End Sub

    Public Property Selected() As Boolean
        Get
            Return boolSelected
        End Get
        Set(ByVal value As Boolean)
            boolSelected = value
        End Set
    End Property

    Public Property GoalNamePlural() As String
        Get
            Return strGoalNamePlural
        End Get
        Set(ByVal value As String)
            strGoalNamePlural = value
        End Set
    End Property

    Public Property GoalName() As String
        Get
            Return strGoalName
        End Get
        Set(ByVal value As String)
            strGoalName = value
        End Set
    End Property

    Public Property PurposeNamePlural() As String
        Get
            Return strPurposeNamePlural
        End Get
        Set(ByVal value As String)
            strPurposeNamePlural = value
        End Set
    End Property

    Public Property PurposeName() As String
        Get
            Return strPurposeName
        End Get
        Set(ByVal value As String)
            strPurposeName = value
        End Set
    End Property

    Public Property OutputNamePlural() As String
        Get
            Return strOutputNamePlural
        End Get
        Set(ByVal value As String)
            strOutputNamePlural = value
        End Set
    End Property

    Public Property OutputName() As String
        Get
            Return strOutputName
        End Get
        Set(ByVal value As String)
            strOutputName = value
        End Set
    End Property

    Public Property ActivityNamePlural() As String
        Get
            Return strActivityNamePlural
        End Get
        Set(ByVal value As String)
            strActivityNamePlural = value
        End Set
    End Property

    Public Property ActivityName() As String
        Get
            Return strActivityName
        End Get
        Set(ByVal value As String)
            strActivityName = value
        End Set
    End Property
End Class
