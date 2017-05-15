Public Class ListViewTargetGroupStatistics
    Inherits ListView

    Private objTargetGroups As TargetGroups

    Public Property TargetGroups() As TargetGroups
        Get
            Return objTargetGroups
        End Get
        Set(ByVal value As TargetGroups)
            objTargetGroups = value
            Me.LoadItems()
        End Set
    End Property


    Public Sub New()
        With Me
            .View = View.Details
            .FullRowSelect = True
            .ForeColor = Color.White
            .BackColor = SystemColors.WindowFrame

            .Columns.Add(LANG_Type, 400, HorizontalAlignment.Left)
            .Columns.Add(LANG_Number, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberFemales, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberMales, 100, HorizontalAlignment.Right)
            .Columns.Add(LANG_NumberPeople, 100, HorizontalAlignment.Right)
        End With
    End Sub

    Public Sub LoadItems()
        Dim intNumber(8), intTotalNumber As Integer
        Dim intNumberOfMales(8), intTotalNumberOfMales, intNumberOfFemales(8), intTotalNumberOfFemales As Integer
        Dim intNumberOfPeople(8), intTotalNumberOfPeople As Integer
        Dim newItem As ListViewItem

        If Me.TargetGroups IsNot Nothing Then
            Me.Items.Clear()
            For Each selTargetGroup As TargetGroup In Me.TargetGroups
                intNumber(selTargetGroup.Type) += selTargetGroup.Number
                intNumberOfMales(selTargetGroup.Type) += selTargetGroup.NumberOfMales
                intNumberOfFemales(selTargetGroup.Type) += selTargetGroup.NumberOfFemales
                intNumberOfPeople(selTargetGroup.Type) += selTargetGroup.NumberOfPeople
            Next

            For i = 0 To 8
                If intNumber(i) > 0 Or intNumberOfMales(i) > 0 Or intNumberOfFemales(i) > 0 Or intNumberOfPeople(i) > 0 Then
                    intTotalNumber += intNumber(i)
                    intTotalNumberOfFemales += intNumberOfFemales(i)
                    intTotalNumberOfMales += intNumberOfMales(i)
                    intTotalNumberOfPeople += intNumberOfPeople(i)

                    newItem = New ListViewItem(LIST_TargetGroupTypes(i))

                    newItem.SubItems.Add(intNumber(i))
                    newItem.SubItems.Add(intNumberOfFemales(i))
                    newItem.SubItems.Add(intNumberOfMales(i))
                    newItem.SubItems.Add(intNumberOfPeople(i))

                    Me.Items.Add(newItem)
                End If
            Next

            newItem = New ListViewItem(LANG_Total)

            newItem.SubItems.Add(intTotalNumber)
            newItem.SubItems.Add(intTotalNumberOfFemales)
            newItem.SubItems.Add(intTotalNumberOfMales)
            newItem.SubItems.Add(intTotalNumberOfPeople)
            newItem.Font = New Font(newItem.Font, FontStyle.Bold)
            Me.Items.Add(newItem)
        End If
    End Sub
End Class
