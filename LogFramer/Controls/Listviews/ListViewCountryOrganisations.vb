Public Class ListViewCountryOrganisations
    Inherits ListViewSortable

    Private objOrganisations As New Organisations

    Public Property Organisations() As Organisations
        Get
            Return objOrganisations
        End Get
        Set(ByVal value As Organisations)
            objOrganisations = value
            Me.LoadItems()
        End Set
    End Property

    Public ReadOnly Property SelectedIdOrganisation() As Integer
        Get
            If SelectedItems.Count > 0 Then
                Dim selItem As ListViewItem = SelectedItems(0)
                Dim intId As Integer

                Integer.TryParse(selItem.Name, intId)
                Return intId
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedIdOrganisations() As Integer()
        Get
            If SelectedItems.Count > 0 Then
                Dim intCount As Integer = SelectedItems.Count - 1
                Dim intIdOrganisations(intCount) As Integer
                Dim selItem As ListViewItem
                Dim intId As Integer

                For i = 0 To intCount
                    selItem = SelectedItems(i)
                    Integer.TryParse(selItem.Name, intId)

                    intIdOrganisations(i) = intId
                Next
                Return intIdOrganisations
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .Columns.Add(LANG_Name, 250, HorizontalAlignment.Left)
            .Columns.Add(LANG_Acronym, 120, HorizontalAlignment.Left)
        End With
    End Sub

    Public ReadOnly Property SelectedOrganisation() As Organisation
        Get
            If Me.SelectedItems.Count > 0 Then
                Dim selOrganisation As Organisation = GetOrganisationById(Me.SelectedIdOrganisation)
                Return selOrganisation
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property SelectedOrganisations() As Organisation()
        Get
            Dim objOrganisations(SelectedItems.Count - 1) As Organisation

            If Me.SelectedGuids.Count > 0 Then
                For i = 0 To SelectedGuids.Count - 1
                    objOrganisations(i) = GetOrganisationById(Me.SelectedIdOrganisations(i))
                Next

                Return objOrganisations
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Function GetOrganisationById(ByVal intIdOrganisation As Integer) As Organisation
        Dim selOrganisation As Organisation = Nothing
        For Each objOrganisation As Organisation In Organisations
            If objOrganisation.idOrganisation = intIdOrganisation Then
                selOrganisation = objOrganisation
                Exit For
            End If
        Next
        Return selOrganisation
    End Function

    Public Sub LoadItems()
        Me.Items.Clear()

        If Me.Organisations IsNot Nothing Then
            For Each selOrganisation As Organisation In Me.Organisations
                Dim newItem As New ListViewItem(selOrganisation.Name)
                newItem.Name = selOrganisation.idOrganisation.ToString
                newItem.SubItems.Add(selOrganisation.Acronym)
                Me.Items.Add(newItem)
            Next
        End If
        Me.SortColumnIndex = 0
        Me.Sort()
    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        MyBase.OnDoubleClick(e)
    End Sub

    Public Overrides Sub NewItem()

    End Sub

    Public Overrides Sub EditItem()

    End Sub

    Public Overrides Sub RemoveItem()

    End Sub

    Public Overrides Sub CutItems()

    End Sub

    Public Overrides Sub CopyItems()

    End Sub

    Public Overrides Sub PasteItems(ByVal PasteItems As ClipboardItems)

    End Sub
End Class

