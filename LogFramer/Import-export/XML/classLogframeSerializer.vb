Imports System.Xml.Serialization
Imports System.Reflection

Public Class LogframeSerializer
    Inherits XmlSerializer


    Public Sub New(ByVal type As Type)
        MyBase.New(type)
    End Sub

    Private Sub LogframeSerializer_UnknownElement(ByVal sender As Object, ByVal e As System.Xml.Serialization.XmlElementEventArgs) Handles Me.UnknownElement
        'Handle changes in the XML schema
        Dim strProperty As String = e.Element.Name

        Select Case strProperty
            Case "Type", "Organisation", "Location", "Relative", "StartDate", "Period", "PeriodUnit", "PeriodDirection", "GuidReferenceMoment", _
                "Duration", "DurationUnit", "DurationUntilEnd", "Repeat", "RepeatUnit", "RepeatTimes", "RepeatUntilEnd", _
                "Preparation", "PreparationUnit", "PreparationFromStart", "PreparationRepeat", _
                "FollowUp", "FollowUpUnit", "FollowUpUntilEnd", "FollowUpRepeat"

                Dim selActivity = TryCast(e.ObjectBeingDeserialized, Activity)
                Dim selValue As String = e.Element.InnerText
                If selActivity IsNot Nothing And String.IsNullOrEmpty(selValue) = False Then
                    Dim selActivityDetail As ActivityDetail = selActivity.ActivityDetail
                    Dim PropInfo As System.Reflection.PropertyInfo = GetType(ActivityDetail).GetProperty(strProperty)

                    If PropInfo IsNot Nothing And selActivityDetail IsNot Nothing Then

                        If PropInfo.PropertyType Is GetType(Int32) Then
                            PropInfo.SetValue(selActivityDetail, Integer.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Single) Then
                            PropInfo.SetValue(selActivityDetail, Single.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Boolean) Then
                            PropInfo.SetValue(selActivityDetail, Boolean.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(String) Then
                            PropInfo.SetValue(selActivityDetail, selValue.ToString, Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Date) Then
                            PropInfo.SetValue(selActivityDetail, Date.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Guid) Then
                            PropInfo.SetValue(selActivityDetail, Guid.Parse(selValue), Nothing)
                        Else
                            PropInfo.SetValue(selActivityDetail, selValue, Nothing)
                        End If

                    End If
                    'selActivity.ActivityDetail.Duration = e.Element.InnerText

                End If
            Case "Repeat"
                Dim selActivity = TryCast(e.ObjectBeingDeserialized, Activity)
                Dim strRepeat As String = e.Element.InnerText
                If selActivity IsNot Nothing And String.IsNullOrEmpty(strRepeat) = False Then
                    Dim selActivityDetail As ActivityDetail = selActivity.ActivityDetail
                    selActivityDetail.RepeatEvery = Single.Parse(strRepeat)
                End If
            Case "RepeatStartDates"
                Dim selActivity = TryCast(e.ObjectBeingDeserialized, Activity)

                If selActivity IsNot Nothing Then
                    For i = 0 To e.Element.ChildNodes.Count - 1
                        Dim strDate As String = e.Element.ChildNodes(i).InnerText
                        Dim selDate As Date
                        If Date.TryParse(strDate, selDate) Then
                            selActivity.ActivityDetail.RepeatStartDates.Add(selDate)
                        End If
                    Next i
                End If
            Case "RepeatEndDates"
                Dim selActivity = TryCast(e.ObjectBeingDeserialized, Activity)

                If selActivity IsNot Nothing Then
                    For i = 0 To e.Element.ChildNodes.Count - 1
                        Dim strDate As String = e.Element.ChildNodes(i).InnerText
                        Dim selDate As Date
                        If Date.TryParse(strDate, selDate) Then
                            selActivity.ActivityDetail.RepeatEndDates.Add(selDate)
                        End If
                    Next i
                End If
            Case "ReferenceMoment"
                Dim selKeyMoment As KeyMoment = TryCast(e.ObjectBeingDeserialized, KeyMoment)
                Dim strReferenceMoment As String = e.Element.InnerText
                If selKeyMoment IsNot Nothing And String.IsNullOrEmpty(strReferenceMoment) = False Then
                    selKeyMoment.GuidReferenceMoment = Guid.Parse(strReferenceMoment)
                End If
            Case "RiskCategory", "Likelihood", "RiskImpact", "RiskResponse"
                Dim selAssumption = TryCast(e.ObjectBeingDeserialized, Assumption)
                Dim selValue As String = e.Element.InnerText

                If selAssumption IsNot Nothing And String.IsNullOrEmpty(selValue) = False Then
                    Dim selRiskDetail As RiskDetail = selAssumption.RiskDetail
                    Dim PropInfo As System.Reflection.PropertyInfo = GetType(RiskDetail).GetProperty(strProperty)

                    If PropInfo IsNot Nothing And selRiskDetail IsNot Nothing Then
                        If PropInfo.PropertyType Is GetType(Int32) Then
                            PropInfo.SetValue(selRiskDetail, Integer.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Single) Then
                            PropInfo.SetValue(selRiskDetail, Single.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Boolean) Then
                            PropInfo.SetValue(selRiskDetail, Boolean.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(String) Then
                            PropInfo.SetValue(selRiskDetail, selValue.ToString, Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Date) Then
                            PropInfo.SetValue(selRiskDetail, Date.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Guid) Then
                            PropInfo.SetValue(selRiskDetail, Guid.Parse(selValue), Nothing)
                        Else
                            PropInfo.SetValue(selRiskDetail, selValue, Nothing)
                        End If
                    End If
                End If
            Case "RiskOwner"
                Dim selAssumption As Assumption = TryCast(e.ObjectBeingDeserialized, Assumption)
                Dim strOwner As String = e.Element.InnerText
                If selAssumption IsNot Nothing And String.IsNullOrEmpty(strOwner) = False Then
                    selAssumption.Owner = strOwner
                End If
            Case "NrDecimals", "Unit", "ValueName"
                Dim selStatement = TryCast(e.ObjectBeingDeserialized, Statement)
                Dim selValue As String = e.Element.InnerText
                If selStatement IsNot Nothing And String.IsNullOrEmpty(selValue) = False Then
                    If selStatement.ValuesDetail Is Nothing Then selStatement.ValuesDetail = New ValuesDetail
                    Dim selValuesDetail As ValuesDetail = selStatement.ValuesDetail
                    Dim PropInfo As System.Reflection.PropertyInfo = GetType(ValuesDetail).GetProperty(strProperty)

                    If PropInfo IsNot Nothing And selValuesDetail IsNot Nothing Then

                        If PropInfo.PropertyType Is GetType(Int32) Then
                            PropInfo.SetValue(selValuesDetail, Integer.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Single) Then
                            PropInfo.SetValue(selValuesDetail, Single.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(Boolean) Then
                            PropInfo.SetValue(selValuesDetail, Boolean.Parse(selValue), Nothing)
                        ElseIf PropInfo.PropertyType Is GetType(String) Then
                            PropInfo.SetValue(selValuesDetail, selValue.ToString, Nothing)
                        Else
                            PropInfo.SetValue(selValuesDetail, selValue, Nothing)
                        End If
                    End If
                End If
            Case "BudgetValue"
                Dim selResource = TryCast(e.ObjectBeingDeserialized, Resource)
                Dim selValue As String = e.Element.InnerText

                If selResource IsNot Nothing And String.IsNullOrEmpty(selValue) = False Then
                    selResource.TotalCostAmount = selValue
                End If
            Case "Abbreviation"
                Dim selOrganisation As Organisation = TryCast(e.ObjectBeingDeserialized, Organisation)
                Dim strAcronym As String = e.Element.InnerText
                If selOrganisation IsNot Nothing And String.IsNullOrEmpty(strAcronym) = False Then
                    selOrganisation.Acronym = strAcronym
                End If
        End Select

    End Sub
End Class
