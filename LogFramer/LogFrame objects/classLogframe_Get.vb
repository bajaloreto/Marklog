Partial Public Class LogFrame

#Region "General Get methods"
    Public Function GetObjectType(ByVal strTypeName As String) As Integer
        Dim intObjectType As Integer = -1
        If [Enum].IsDefined(GetType(ObjectTypes), strTypeName) = True Then

            intObjectType = [Enum].Parse(GetType(ObjectTypes), strTypeName)
        End If
        Return intObjectType
    End Function

    Public Function GetStructSortNumber(ByVal selStruct As Struct) As String
        Dim strSort As String = String.Empty
        Dim strParentSort As String = String.Empty

        Dim intIndex As Integer
        If selStruct Is Nothing Then Return String.Empty

        If selStruct.GetType Is GetType(Goal) Then
            Dim CurrentGoal As Goal = CType(selStruct, Goal)
            intIndex = Me.Goals.IndexOf(CurrentGoal)

            strSort = CreateSortNumber(intIndex)
        ElseIf selStruct.GetType Is GetType(Purpose) Then
            Dim CurrentPurpose As Purpose = CType(selStruct, Purpose)
            intIndex = Me.Purposes.IndexOf(CurrentPurpose)

            strSort = CreateSortNumber(intIndex)
        ElseIf selStruct.GetType Is GetType(Output) Then
            Dim CurrentOutput As Output = CType(selStruct, Output)
            Dim ParentPurpose As Purpose = GetParent(CurrentOutput)

            If ParentPurpose Is Nothing Then Return String.Empty

            strParentSort = GetStructSortNumber(ParentPurpose)
            intIndex = ParentPurpose.Outputs.IndexOf(CurrentOutput)

            strSort = CreateSortNumber(intIndex, strParentSort)
        ElseIf selStruct.GetType Is GetType(Activity) Then
            Dim CurrentActivity As Activity = CType(selStruct, Activity)
            Dim objParent As Object = GetParent(CurrentActivity)
            Dim objActivities As Activities = GetParentCollection(CurrentActivity)

            If objParent Is Nothing Or objActivities Is Nothing Then Return String.Empty

            strParentSort = GetStructSortNumber(objParent)
            intIndex = objActivities.IndexOf(CurrentActivity)

            strSort = CreateSortNumber(intIndex, strParentSort)
        End If

        Return strSort
    End Function

    Public Function GetTargetSystem(ByVal intSection As Integer) As Integer
        Dim objTargetDeadlinesSection As TargetDeadlinesSection = GetTargetDeadlinesSection(intSection)

        If objTargetDeadlinesSection IsNot Nothing Then
            Return objTargetDeadlinesSection.Repetition
        Else
            Return -1
        End If
    End Function

    Public Function GetTargetGroups(Optional ByVal strTargetGroupName As String = "") As TargetGroups
        Dim objTargetGroups As New TargetGroups

        If String.IsNullOrEmpty(strTargetGroupName) Or strTargetGroupName.StartsWith("- ") Then
            For Each selPurpose As Purpose In Me.Purposes
                For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                    objTargetGroups.Add(selTargetGroup)
                Next
            Next
        Else
            For Each selPurpose As Purpose In Me.Purposes
                For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                    If selTargetGroup.Name = strTargetGroupName Then
                        objTargetGroups.Add(selTargetGroup)

                        Exit For
                    End If
                Next
            Next
        End If

        Return objTargetGroups
    End Function
#End Region

#Region "Get by GUID"
    Public Function GetStructByGuid(ByVal objGuid As Guid) As Struct
        Dim selStruct As Struct = GetGoalByGuid(objGuid)

        If selStruct Is Nothing Then selStruct = GetPurposeByGuid(objGuid)
        If selStruct Is Nothing Then selStruct = GetOutputByGuid(objGuid)
        If selStruct Is Nothing Then selStruct = GetActivityByGuid(objGuid)

        Return selStruct
    End Function

    Public Function GetGoalByGuid(ByVal objGuid As Guid) As Goal
        Dim selGoal As Goal = Me.Goals.GetGoalByGuid(objGuid)
        Return selGoal
    End Function

    Public Function GetPurposeByGuid(ByVal objGuid As Guid) As Purpose
        Dim selPurpose As Purpose = Me.Purposes.GetPurposeByGuid(objGuid)
        Return selPurpose
    End Function

    Public Function GetOutputByGuid(ByVal objGuid As Guid) As Output
        Dim selOutput As Output = Nothing
        For Each selPurpose As Purpose In Me.Purposes
            Dim objOutput As Output = selPurpose.Outputs.GetOutputByGuid(objGuid)
            If objOutput IsNot Nothing Then
                selOutput = objOutput
                Exit For
            End If
        Next
        Return selOutput
    End Function

    Public Function GetActivityByGuid(ByVal objGuid As Guid) As Activity
        Dim selActivity As Activity = Nothing
        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                Dim objActivity As Activity = selOutput.Activities.GetActivityByGuid(objGuid)
                If objActivity IsNot Nothing Then
                    selActivity = objActivity
                    Return selActivity
                End If
            Next
        Next
        Return selActivity
    End Function

    Public Function GetIndicatorByGuid(ByVal objGuid As Guid) As Indicator
        Dim selIndicator As Indicator = Nothing
        selIndicator = Me.Goals.GetIndicatorByGuid(objGuid)
        If selIndicator IsNot Nothing Then
            Return selIndicator
        Else
            selIndicator = Me.Purposes.GetIndicatorByGuid(objGuid)
            If selIndicator IsNot Nothing Then
                Return selIndicator
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selIndicator = selPurpose.Outputs.GetIndicatorByGuid(objGuid)
                    If selIndicator IsNot Nothing Then
                        Return selIndicator
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selIndicator = selOutput.Activities.GetIndicatorByGuid(objGuid)
                            If selIndicator IsNot Nothing Then
                                Return selIndicator
                            End If
                        Next
                    End If

                Next
            End If
        End If

        Return selIndicator
    End Function

    Public Function GetStatementByGuid(ByVal objGuid As Guid) As Statement
        Dim selStatement As Statement = Nothing
        selStatement = Me.Goals.GetStatementByGuid(objGuid)
        If selStatement IsNot Nothing Then
            Return selStatement
        Else
            selStatement = Me.Purposes.GetStatementByGuid(objGuid)
            If selStatement IsNot Nothing Then
                Return selStatement
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selStatement = selPurpose.Outputs.GetStatementByGuid(objGuid)
                    If selStatement IsNot Nothing Then
                        Return selStatement
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selStatement = selOutput.Activities.GetStatementByGuid(objGuid)
                            If selStatement IsNot Nothing Then Return selStatement
                        Next
                    End If

                Next
            End If
        End If

        Return selStatement
    End Function

    Public Function GetTargetByGuid(ByVal objGuid As Guid) As Target
        Dim selTarget As Target = Nothing
        selTarget = Me.Goals.GetTargetByGuid(objGuid)
        If selTarget IsNot Nothing Then
            Return selTarget
        Else
            selTarget = Me.Purposes.GetTargetByGuid(objGuid)
            If selTarget IsNot Nothing Then
                Return selTarget
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selTarget = selPurpose.Outputs.GetTargetByGuid(objGuid)
                    If selTarget IsNot Nothing Then
                        Return selTarget
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selTarget = selOutput.Activities.GetTargetByGuid(objGuid)
                            If selTarget IsNot Nothing Then Return selTarget
                        Next
                    End If

                Next
            End If
        End If

        Return selTarget
    End Function

    Public Function GetTargetDeadlineByGuid(ByVal objGuid As Guid) As TargetDeadline
        Dim selTargetDeadline As TargetDeadline = Nothing

        selTargetDeadline = Me.TargetDeadlinesActivities.GetTargetDeadlineByGuid(objGuid)
        If selTargetDeadline IsNot Nothing Then
            Return selTargetDeadline
        Else
            selTargetDeadline = Me.TargetDeadlinesOutputs.GetTargetDeadlineByGuid(objGuid)
            If selTargetDeadline IsNot Nothing Then
                Return selTargetDeadline
            Else
                selTargetDeadline = Me.TargetDeadlinesPurposes.GetTargetDeadlineByGuid(objGuid)
                If selTargetDeadline IsNot Nothing Then
                    Return selTargetDeadline
                Else
                    selTargetDeadline = Me.TargetDeadlinesGoals.GetTargetDeadlineByGuid(objGuid)
                End If
            End If
        End If
        Return selTargetDeadline
    End Function

    Public Function GetResponseClassByGuid(ByVal objGuid As Guid) As ResponseClass
        Dim selResponseClass As ResponseClass = Nothing
        selResponseClass = Me.Goals.GetResponseClassByGuid(objGuid)
        If selResponseClass IsNot Nothing Then
            Return selResponseClass
        Else
            selResponseClass = Me.Purposes.GetResponseClassByGuid(objGuid)
            If selResponseClass IsNot Nothing Then
                Return selResponseClass
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selResponseClass = selPurpose.Outputs.GetResponseClassByGuid(objGuid)
                    If selResponseClass IsNot Nothing Then
                        Return selResponseClass
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selResponseClass = selOutput.Activities.GetResponseClassByGuid(objGuid)
                            If selResponseClass IsNot Nothing Then Return selResponseClass
                        Next
                    End If
                Next
            End If
        End If

        Return selResponseClass
    End Function

    Public Function GetResourceByGuid(ByVal objGuid As Guid) As Resource
        Dim selResource As Resource = Nothing
        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                selResource = selOutput.Activities.GetResourceByGuid(objGuid)
                If selResource IsNot Nothing Then Return selResource
            Next
        Next

        Return selResource
    End Function

    Public Function GetVerificationSourceByGuid(ByVal objGuid As Guid) As VerificationSource
        Dim selVerificationSource As VerificationSource = Nothing
        selVerificationSource = Me.Goals.GetVerificationSourceByGuid(objGuid)
        If selVerificationSource IsNot Nothing Then
            Return selVerificationSource
        Else
            selVerificationSource = Me.Purposes.GetVerificationSourceByGuid(objGuid)
            If selVerificationSource IsNot Nothing Then
                Return selVerificationSource
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selVerificationSource = selPurpose.Outputs.GetVerificationSourceByGuid(objGuid)
                    If selVerificationSource IsNot Nothing Then
                        Return selVerificationSource
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selVerificationSource = selOutput.Activities.GetVerificationSourceByGuid(objGuid)
                            If selVerificationSource IsNot Nothing Then Return selVerificationSource
                        Next
                    End If

                Next
            End If
        End If

        Return selVerificationSource
    End Function

    Public Function GetAssumptionByGuid(ByVal objGuid As Guid) As Assumption
        Dim selAssumption As Assumption = Nothing
        selAssumption = Me.Goals.GetAssumptionByGuid(objGuid)
        If selAssumption IsNot Nothing Then
            Return selAssumption
        Else
            selAssumption = Me.Purposes.GetAssumptionByGuid(objGuid)
            If selAssumption IsNot Nothing Then
                Return selAssumption
            Else
                For Each selPurpose As Purpose In Me.Purposes
                    selAssumption = selPurpose.Outputs.GetAssumptionByGuid(objGuid)
                    If selAssumption IsNot Nothing Then
                        Return selAssumption
                    Else
                        For Each selOutput As Output In selPurpose.Outputs
                            selAssumption = selOutput.Activities.GetAssumptionByGuid(objGuid)
                            If selAssumption IsNot Nothing Then Return selAssumption
                        Next
                    End If

                Next
            End If
        End If

        Return selAssumption
    End Function

    Public Function GetBudgetYearByGuid(ByVal objGuid As Guid) As BudgetYear
        Dim selBudgetYear As BudgetYear = Me.Budget.GetBudgetYearByGuid(objGuid)

        Return selBudgetYear
    End Function

    Public Function GetBudgetItemByGuid(ByVal objGuid As Guid) As BudgetItem
        Dim selBudgetItem As BudgetItem = Me.Budget.GetBudgetItemByGuid(objGuid)

        Return selBudgetItem
    End Function

    Public Function GetProjectPartnerByGuid(ByVal objGuid As Guid) As ProjectPartner
        Dim selPartner As ProjectPartner = Nothing
        selPartner = Me.ProjectPartners.GetProjectPartnerByGuid(objGuid)
        Return selPartner
    End Function

    Public Function GetProjectPartnerByOrganisationGuid(ByVal objOrganisationGuid As Guid) As ProjectPartner
        Dim selPartner As ProjectPartner = Nothing
        selPartner = Me.ProjectPartners.GetProjectPartnerByOrganisationGuid(objOrganisationGuid)
        Return selPartner
    End Function

    Public Function GetOrganisationByGuid(ByVal objGuid As Guid) As Organisation
        Dim selOrganisation As Organisation = Nothing
        selOrganisation = Me.ProjectPartners.GetOrganisationByGuid(objGuid)
        Return selOrganisation
    End Function

    Public Function GetAddressByGuid(ByVal objGuid As Guid) As Address
        Dim selAddress As Address = Nothing
        For Each selPartner As ProjectPartner In Me.ProjectPartners
            selAddress = selPartner.Organisation.Addresses.GetAddressByGuid(objGuid)
            If selAddress IsNot Nothing Then
                Return selAddress
            Else
                For Each selContact As Contact In selPartner.Organisation.Contacts
                    selAddress = selContact.Addresses.GetAddressByGuid(objGuid)
                    If selAddress IsNot Nothing Then Return selAddress
                Next
            End If
        Next
        Return selAddress
    End Function

    Public Function GetTelephoneNumberByGuid(ByVal objGuid As Guid) As TelephoneNumber
        Dim selTelephoneNumber As TelephoneNumber = Nothing
        For Each selPartner As ProjectPartner In Me.ProjectPartners
            selTelephoneNumber = selPartner.Organisation.TelephoneNumbers.GetTelephoneNumberByGuid(objGuid)
            If selTelephoneNumber IsNot Nothing Then
                Return selTelephoneNumber
            Else
                For Each selContact As Contact In selPartner.Organisation.Contacts
                    selTelephoneNumber = selContact.TelephoneNumbers.GetTelephoneNumberByGuid(objGuid)
                    If selTelephoneNumber IsNot Nothing Then Return selTelephoneNumber
                Next
            End If
        Next
        Return selTelephoneNumber
    End Function

    Public Function GetEmailByGuid(ByVal objGuid As Guid) As Email
        Dim selEmail As Email = Nothing
        For Each selPartner As ProjectPartner In Me.ProjectPartners
            selEmail = selPartner.Organisation.Emails.GetEmailByGuid(objGuid)
            If selEmail IsNot Nothing Then
                Return selEmail
            Else
                For Each selContact As Contact In selPartner.Organisation.Contacts
                    selEmail = selContact.Emails.GetEmailByGuid(objGuid)
                    If selEmail IsNot Nothing Then Return selEmail
                Next
            End If
        Next
        Return selEmail
    End Function

    Public Function GetWebsiteByGuid(ByVal objGuid As Guid) As Website
        Dim selWebsite As Website = Nothing
        For Each selPartner As ProjectPartner In Me.ProjectPartners
            selWebsite = selPartner.Organisation.WebSites.GetWebsiteByGuid(objGuid)
            If selWebsite IsNot Nothing Then
                Return selWebsite
            Else
                For Each selContact As Contact In selPartner.Organisation.Contacts
                    selWebsite = selContact.Websites.GetWebsiteByGuid(objGuid)
                    If selWebsite IsNot Nothing Then Return selWebsite
                Next
            End If
        Next
        Return selWebsite
    End Function

    Public Function GetContactByGuid(ByVal objGuid As Guid) As Contact
        Dim selContact As Contact = Nothing
        For Each selPartner As ProjectPartner In Me.ProjectPartners
            selContact = selPartner.Organisation.Contacts.GetContactByGuid(objGuid)
            If selContact IsNot Nothing Then
                Return selContact
            End If
        Next
        Return selContact
    End Function

    Public Function GetTargetGroupByGuid(ByVal objGuid As Guid) As TargetGroup
        Dim selTargetGroup As TargetGroup = Nothing
        selTargetGroup = Me.Purposes.GetTargetGroupByGuid(objGuid)

        Return selTargetGroup
    End Function

    Public Function GetTargetGroupInformationByGuid(ByVal objGuid As Guid) As TargetGroupInformation
        Dim selTargetGroupInformation As TargetGroupInformation = Nothing
        For Each selPurpose As Purpose In Me.Purposes
            For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                selTargetGroupInformation = selTargetGroup.TargetGroupInformations.GetTargetGroupInformationByGuid(objGuid)
                If selTargetGroupInformation IsNot Nothing Then Return selTargetGroupInformation
            Next
        Next

        Return selTargetGroupInformation
    End Function

    Public Function GetKeyMomentByGuid(ByVal objGuid As Guid) As KeyMoment
        Dim selKeyMoment As KeyMoment = Nothing
        For Each selPurpose As Purpose In Me.Purposes
            selKeyMoment = selPurpose.Outputs.GetKeyMomentByGuid(objGuid)
            If selKeyMoment IsNot Nothing Then Return selKeyMoment
        Next

        Return selKeyMoment
    End Function

    Public Function GetReferenceDateByGuid(ByVal objGuid As Guid, Optional ByVal boolBefore As Boolean = False) As Date
        Dim selDate As Date
        For Each selMoment As KeyMoment In Me.Keymoments
            If selMoment.Guid = objGuid Then
                selDate = selMoment.ExactDateKeyMoment
                Return selDate
            End If
        Next

        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                For Each selMoment As KeyMoment In selOutput.KeyMoments
                    If selMoment.Guid = objGuid Then
                        selDate = selMoment.ExactDateKeyMoment
                        Return selDate
                    End If
                Next
                selDate = GetReferenceDateByGuid_Activities(objGuid, boolBefore, selOutput.Activities)
                If selDate > Date.MinValue Then Return selDate
            Next
        Next
        Return Date.MinValue
    End Function

    Public Function GetReferenceDateByGuid_Activities(ByVal objGuid As Guid, ByVal boolBefore As Boolean, ByVal selActivities As Activities) As Date
        Dim selDate As Date

        For Each selActivity As Activity In selActivities
            If selActivity.Guid = objGuid Then
                If boolBefore = False Then
                    selDate = selActivity.ExactEndDate
                Else
                    selDate = selActivity.ExactStartDate
                End If
                Return selDate
            ElseIf selActivity.Activities.Count > 0 Then
                selDate = GetReferenceDateByGuid_Activities(objGuid, boolBefore, selActivity.Activities)
                If selDate > Date.MinValue Then Return selDate
            End If
        Next

        Return Date.MinValue
    End Function

    Public Function GetReferenceMomentByGuid(ByVal objGuid As Guid) As Object
        Dim selMoment As Object = Nothing

        selMoment = GetReferenceMomentByGuid_KeyMoments(objGuid, Me.Keymoments)

        If selMoment Is Nothing Then
            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    selMoment = GetReferenceMomentByGuid_KeyMoments(objGuid, selOutput.KeyMoments)
                    If selMoment IsNot Nothing Then Exit For

                    selMoment = GetReferenceMomentByGuid_Activities(objGuid, selOutput.Activities)
                    If selMoment IsNot Nothing Then Exit For
                Next
            Next
        End If

        Return selMoment
    End Function

    Public Function GetReferenceMomentByGuid_KeyMoments(ByVal objGuid As Guid, ByVal selKeyMoments As KeyMoments) As Object
        For Each selKeyMoment As KeyMoment In selKeyMoments
            If selKeyMoment.Guid = objGuid Then
                Return selKeyMoment
            End If
        Next

        Return Nothing
    End Function

    Public Function GetReferenceMomentByGuid_Activities(ByVal objGuid As Guid, ByVal selActivities As Activities) As Object
        Dim selMoment As Object = Nothing

        For Each selActivity As Activity In selActivities
            If selActivity.Guid = objGuid Then

                selMoment = selActivity
                Exit For
            End If
            If selActivity.Activities.Count > 0 Then
                selMoment = GetReferenceMomentByGuid_Activities(objGuid, selActivity.Activities)
                If selMoment IsNot Nothing Then Exit For
            End If
        Next

        Return selMoment
    End Function

    Public Function GetReferingMomentsByReferenceGuid(ByVal objReferenceGuid As Guid) As ArrayList
        Dim ReferersList As New ArrayList

        ReferersList = GetReferingMomentsByReferenceGuid_KeyMoments(ReferersList, objReferenceGuid, Me.Keymoments)

        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                ReferersList = GetReferingMomentsByReferenceGuid_KeyMoments(ReferersList, objReferenceGuid, selOutput.KeyMoments)

                ReferersList = GetReferingMomentsByReferenceGuid_Activities(ReferersList, objReferenceGuid, selOutput.Activities)
            Next
        Next

        Return ReferersList
    End Function

    Public Function GetReferingMomentsByReferenceGuid_KeyMoments(ByVal ReferersList As ArrayList, ByVal objGuid As Guid, ByVal selKeyMoments As KeyMoments) As ArrayList
        For Each selKeyMoment As KeyMoment In selKeyMoments
            If selKeyMoment.GuidReferenceMoment = objGuid Then
                ReferersList.Add(selKeyMoment)
            End If
        Next

        Return ReferersList
    End Function

    Public Function GetReferingMomentsByReferenceGuid_Activities(ByVal ReferersList As ArrayList, ByVal objGuid As Guid, ByVal selActivities As Activities) As Object
        For Each selActivity As Activity In selActivities
            If selActivity.ActivityDetail.GuidReferenceMoment = objGuid Then
                ReferersList.Add(selActivity)
            ElseIf selActivity.Activities.Count > 0 Then
                ReferersList = GetReferingMomentsByReferenceGuid_Activities(ReferersList, objGuid, selActivity.Activities)
            End If
        Next

        Return ReferersList
    End Function

    Public Function GetReferingResourcesByReferenceGuid(ByVal objGuid As Guid) As List(Of BudgetItemReference)
        Dim objReferences As New List(Of BudgetItemReference)

        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                objReferences = GetReferingResourcesByReferenceGuid_Activities(objReferences, objGuid, selOutput.Activities)
            Next
        Next

        Return objReferences
    End Function

    Public Function GetReferingResourcesByReferenceGuid_Activities(ByVal objReferences As List(Of BudgetItemReference), ByVal objGuid As Guid, ByVal objActivities As Activities) As List(Of BudgetItemReference)
        For Each selActivity As Activity In objActivities

            For Each selResource As Resource In selActivity.Resources
                For Each selBudgetItemReference As BudgetItemReference In selResource.BudgetItemReferences
                    If selBudgetItemReference.ReferenceBudgetItemGuid = objGuid Then _
                        objReferences.Add(selBudgetItemReference)
                Next
            Next

            If selActivity.Activities.Count > 0 Then
                objReferences = GetReferingResourcesByReferenceGuid_Activities(objReferences, objGuid, selActivity.Activities)
            End If
        Next

        Return objReferences
    End Function
#End Region

#Region "Get parents"
    Public Function GetParent(ByVal selItem As Object) As Object
        Dim ParentItem As Object = Nothing

        If selItem Is Nothing Then Return Nothing

        If TryCast(selItem, LogframeObject) IsNot Nothing Then
            ParentItem = GetParent_LogframeObject(selItem)
        Else
            ParentItem = GetParent_ProjectObject(selItem)
        End If

        Return ParentItem
    End Function

    Public Function GetParent_LogframeObject(ByVal selItem As Object) As Object
        Dim ParentItem As Object = Nothing

        If selItem Is Nothing Then Return Nothing

        Select Case selItem.GetType
            Case GetType(Goal)
                ParentItem = Me
            Case GetType(Purpose)
                ParentItem = Me
            Case GetType(Output)
                Dim selOutput As Output = DirectCast(selItem, Output)
                ParentItem = CurrentLogFrame.GetPurposeByGuid(selOutput.ParentPurposeGuid)
            Case GetType(Activity)
                Dim selActivity As Activity = DirectCast(selItem, Activity)
                If selActivity.ParentOutputGuid <> Guid.Empty Then
                    ParentItem = CurrentLogFrame.GetOutputByGuid(selActivity.ParentOutputGuid)
                Else
                    ParentItem = CurrentLogFrame.GetActivityByGuid(selActivity.ParentActivityGuid)
                End If
            Case GetType(Indicator)
                Dim selIndicator As Indicator = DirectCast(selItem, Indicator)
                If selIndicator.ParentStructGuid <> Guid.Empty Then
                    ParentItem = CurrentLogFrame.GetStructByGuid(selIndicator.ParentStructGuid)
                Else
                    ParentItem = CurrentLogFrame.GetIndicatorByGuid(selIndicator.ParentIndicatorGuid)
                End If
            Case GetType(Statement)
                Dim selStatement As Statement = DirectCast(selItem, Statement)
                ParentItem = CurrentLogFrame.GetIndicatorByGuid(selStatement.ParentIndicatorGuid)
            Case GetType(Resource)
                Dim selResource As Resource = DirectCast(selItem, Resource)
                ParentItem = CurrentLogFrame.GetActivityByGuid(selResource.ParentStructGuid)
            Case GetType(VerificationSource)
                Dim selVerificationSource As VerificationSource = DirectCast(selItem, VerificationSource)
                ParentItem = CurrentLogFrame.GetIndicatorByGuid(selVerificationSource.ParentIndicatorGuid)
            Case GetType(Assumption)
                Dim selAssumption As Assumption = DirectCast(selItem, Assumption)
                ParentItem = CurrentLogFrame.GetStructByGuid(selAssumption.ParentStructGuid)
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = DirectCast(selItem, BudgetItem)
                If selBudgetItem.ParentBudgetYearGuid <> Guid.Empty Then
                    ParentItem = CurrentLogFrame.GetBudgetYearByGuid(selBudgetItem.ParentBudgetYearGuid)
                Else
                    ParentItem = CurrentLogFrame.GetBudgetItemByGuid(selBudgetItem.ParentBudgetItemGuid)
                End If
        End Select

        Return ParentItem
    End Function

    Public Function GetParent_ProjectObject(ByVal selItem As Object) As Object
        Dim ParentItem As Object = Nothing

        If selItem Is Nothing Then Return Nothing

        Select Case selItem.GetType
            Case GetType(Address)
                Dim selAddress As Address = DirectCast(selItem, Address)
                ParentItem = GetParent_OrganisationContact(selAddress.ParentGuid)
            Case GetType(Contact)
                Dim selContact As Contact = DirectCast(selItem, Contact)
                ParentItem = GetOrganisationByGuid(selContact.ParentOrganisationGuid)
            Case GetType(Email)
                Dim selEmail As Email = DirectCast(selItem, Email)
                ParentItem = GetParent_OrganisationContact(selEmail.ParentGuid)
            Case GetType(KeyMoment)
                Dim selKeyMoment As KeyMoment = DirectCast(selItem, KeyMoment)
                ParentItem = GetOutputByGuid(selKeyMoment.ParentOutputGuid)
            Case GetType(LogFrame)
                'do nothing
            Case GetType(ProjectPartner)
                ParentItem = Me
            Case GetType(TargetGroup)
                Dim selTargetGroup As TargetGroup = DirectCast(selItem, TargetGroup)
                ParentItem = GetPurposeByGuid(selTargetGroup.ParentPurposeGuid)
            Case GetType(TargetGroupInformation)
                Dim selTargetGroupInfo As TargetGroupInformation = DirectCast(selItem, TargetGroupInformation)
                ParentItem = GetTargetGroupByGuid(selTargetGroupInfo.ParentTargetGroupGuid)
            Case GetType(TelephoneNumber)
                Dim selTelephoneNumber As TelephoneNumber = DirectCast(selItem, TelephoneNumber)
                ParentItem = GetParent_OrganisationContact(selTelephoneNumber.ParentGuid)
            Case GetType(Website)
                Dim selWebsite As Website = DirectCast(selItem, Website)
                ParentItem = GetParent_OrganisationContact(selWebsite.ParentGuid)
        End Select

        Return ParentItem
    End Function

    Public Function GetParent_OrganisationContact(ByVal selParentGuid As Guid)
        Dim ParentItem As Object = Nothing
        Dim selOrganisation As Organisation = GetOrganisationByGuid(selParentGuid)

        If selOrganisation IsNot Nothing Then
            ParentItem = selOrganisation
        Else
            Dim selContact As Contact = GetContactByGuid(selParentGuid)

            If selContact IsNot Nothing Then
                ParentItem = selContact
            End If
        End If

        Return ParentItem
    End Function

    Public Function GetParentCollection(ByVal selItem As Object) As Object
        Dim ParentCollection As Object = Nothing

        If selItem Is Nothing Then Return Nothing

        If TryCast(selItem, LogframeObject) IsNot Nothing Then
            ParentCollection = GetParentCollection_LogframeObject(selItem)
        Else
            ParentCollection = GetParentCollection_ProjectObject(selItem)
        End If

        Return ParentCollection
    End Function

    Public Function GetParentCollection_LogframeObject(ByVal selItem As Object) As Object
        Dim ParentCollection As Object = Nothing

        Select Case selItem.GetType
            Case GetType(Goal)
                ParentCollection = Me.Goals
            Case GetType(Purpose)
                ParentCollection = Me.Purposes
            Case GetType(Output)
                Dim selOutput As Output = DirectCast(selItem, Output)
                Dim ParentPurpose As Purpose = GetPurposeByGuid(selOutput.ParentPurposeGuid)

                If ParentPurpose IsNot Nothing Then ParentCollection = ParentPurpose.Outputs
            Case GetType(Activity)
                Dim selActivity As Activity = DirectCast(selItem, Activity)
                If selActivity.ParentOutputGuid <> Guid.Empty Then
                    Dim ParentOutput As Output = GetOutputByGuid(selActivity.ParentOutputGuid)

                    If ParentOutput IsNot Nothing Then ParentCollection = ParentOutput.Activities
                Else
                    Dim ParentActivity As Activity = GetActivityByGuid(selActivity.ParentActivityGuid)

                    If ParentActivity IsNot Nothing Then ParentCollection = ParentActivity.Activities
                End If
            Case GetType(Indicator)
                Dim selIndicator As Indicator = DirectCast(selItem, Indicator)
                If selIndicator.ParentStructGuid <> Guid.Empty Then
                    Dim ParentStruct As Struct = GetStructByGuid(selIndicator.ParentStructGuid)

                    If ParentStruct IsNot Nothing Then ParentCollection = ParentStruct.Indicators
                Else
                    Dim ParentIndicator As Indicator = GetIndicatorByGuid(selIndicator.ParentIndicatorGuid)

                    If ParentIndicator IsNot Nothing Then ParentCollection = ParentIndicator.Indicators
                End If
            Case GetType(Statement)
                Dim selStatement As Statement = DirectCast(selItem, Statement)
                Dim ParentIndicator As Indicator = GetIndicatorByGuid(selStatement.ParentIndicatorGuid)

                If ParentIndicator IsNot Nothing Then ParentCollection = ParentIndicator.Statements
            Case GetType(Resource)
                Dim selResource As Resource = DirectCast(selItem, Resource)
                Dim ParentActivity As Activity = GetActivityByGuid(selResource.ParentStructGuid)

                If ParentActivity IsNot Nothing Then ParentCollection = ParentActivity.Resources
            Case GetType(VerificationSource)
                Dim selVerificationSource As VerificationSource = DirectCast(selItem, VerificationSource)
                Dim ParentIndicator As Indicator = GetIndicatorByGuid(selVerificationSource.ParentIndicatorGuid)

                If ParentIndicator IsNot Nothing Then ParentCollection = ParentIndicator.VerificationSources
            Case GetType(Assumption)
                Dim selAssumption As Assumption = DirectCast(selItem, Assumption)
                Dim ParentStruct As Struct = GetStructByGuid(selAssumption.ParentStructGuid)

                If ParentStruct IsNot Nothing Then ParentCollection = ParentStruct.Assumptions
            Case GetType(BudgetItem)
                Dim selBudgetItem As BudgetItem = DirectCast(selItem, BudgetItem)
                If selBudgetItem.ParentBudgetYearGuid <> Guid.Empty Then
                    Dim ParentBudgetYear As BudgetYear = GetBudgetYearByGuid(selBudgetItem.ParentBudgetYearGuid)

                    If ParentBudgetYear IsNot Nothing Then ParentCollection = ParentBudgetYear.BudgetItems
                Else
                    Dim ParentBudgetItem As BudgetItem = GetBudgetItemByGuid(selBudgetItem.ParentBudgetItemGuid)

                    If ParentBudgetItem IsNot Nothing Then ParentCollection = ParentBudgetItem.BudgetItems
                End If
        End Select

        Return ParentCollection
    End Function

    Public Function GetParentCollection_ProjectObject(ByVal selItem As Object) As Object
        Dim ParentItem As Object = Nothing
        Dim ParentCollection As Object = Nothing

        Select Case selItem.GetType
            Case GetType(Address)
                Dim selAddress As Address = DirectCast(selItem, Address)
                ParentItem = GetParent_OrganisationContact(selAddress.ParentGuid)
                If ParentItem IsNot Nothing Then ParentCollection = ParentItem.Addresses
            Case GetType(Contact)
                Dim selContact As Contact = DirectCast(selItem, Contact)
                Dim ParentOrganisation As Organisation = GetOrganisationByGuid(selContact.ParentOrganisationGuid)

                ParentCollection = ParentOrganisation.Contacts
            Case GetType(Email)
                Dim selEmail As Email = DirectCast(selItem, Email)
                ParentItem = GetParent_OrganisationContact(selEmail.ParentGuid)
                If ParentItem IsNot Nothing Then ParentCollection = ParentItem.Emails
            Case GetType(KeyMoment)
                Dim selKeyMoment As KeyMoment = DirectCast(selItem, KeyMoment)
                Dim ParentOutput As Output = GetOutputByGuid(selKeyMoment.ParentOutputGuid)

                If ParentOutput IsNot Nothing Then ParentCollection = ParentOutput.KeyMoments
            Case GetType(LogFrame)
                'do nothing
            Case GetType(ProjectPartner)
                ParentCollection = Me.ProjectPartners
            Case GetType(TargetGroup)
                Dim selTargetGroup As TargetGroup = DirectCast(selItem, TargetGroup)
                Dim ParentPurpose As Purpose = GetPurposeByGuid(selTargetGroup.ParentPurposeGuid)

                If ParentPurpose IsNot Nothing Then ParentCollection = ParentPurpose.TargetGroups
            Case GetType(TargetGroupInformation)
                Dim selTargetGroupInformation As TargetGroupInformation = DirectCast(selItem, TargetGroupInformation)
                Dim ParentTargetGroup As TargetGroup = GetTargetGroupByGuid(selTargetGroupInformation.ParentTargetGroupGuid)

                If ParentTargetGroup IsNot Nothing Then ParentCollection = ParentTargetGroup.TargetGroupInformations
            Case GetType(TelephoneNumber)
                Dim selTelephoneNumber As TelephoneNumber = DirectCast(selItem, TelephoneNumber)
                ParentItem = GetParent_OrganisationContact(selTelephoneNumber.ParentGuid)
                If ParentItem IsNot Nothing Then ParentCollection = ParentItem.TelephoneNumbers
            Case GetType(Website)
                Dim selWebsite As Website = DirectCast(selItem, Website)
                ParentItem = GetParent_OrganisationContact(selWebsite.ParentGuid)
                If ParentItem IsNot Nothing Then ParentCollection = ParentItem.Websites
        End Select

        Return ParentCollection
    End Function

    Public Function IsParentLineage(ByVal ChildActivity As Activity, ByVal ParentActivity As Activity) As Boolean
        Dim ParentStruct As Struct = GetParent(ChildActivity)
        Dim boolIsParent As Boolean

        If ChildActivity Is Nothing Or ParentActivity Is Nothing Or ParentStruct Is Nothing Then Return False

        If ParentStruct.GetType() Is GetType(Activity) Then
            Dim selActivity As Activity = DirectCast(ParentStruct, Activity)
            If ParentStruct Is ParentActivity Then
                boolIsParent = True
                Return boolIsParent
            Else
                boolIsParent = IsParentLineage(selActivity, ParentActivity)
                Return boolIsParent
            End If
        Else
            Return boolIsParent
        End If
    End Function

    Public Function IsParentLineage(ByVal ChildIndicator As Indicator, ByVal ParentIndicator As Indicator) As Boolean
        Dim objParent As Object = GetParent(ChildIndicator)
        Dim boolIsParent As Boolean

        If ChildIndicator Is Nothing Or ParentIndicator Is Nothing Or objParent Is Nothing Then Return False

        If objParent.GetType() Is GetType(Indicator) Then
            Dim selIndicator As Indicator = DirectCast(objParent, Indicator)
            If objParent Is ParentIndicator Then
                boolIsParent = True
                Return boolIsParent
            Else
                boolIsParent = IsParentLineage(selIndicator, ParentIndicator)
                Return boolIsParent
            End If
        Else
            Return boolIsParent
        End If
    End Function

    Public Function IsParentLineage(ByVal ChildBudgetItem As BudgetItem, ByVal ParentBudgetItem As BudgetItem) As Boolean
        Dim objParent As Object = GetParent(ChildBudgetItem)
        Dim boolIsParent As Boolean

        If ChildBudgetItem Is Nothing Or ParentBudgetItem Is Nothing Or objParent Is Nothing Then Return False

        If objParent.GetType() Is GetType(BudgetItem) Then
            Dim selBudgetItem As BudgetItem = DirectCast(objParent, BudgetItem)
            If objParent Is ParentBudgetItem Then
                boolIsParent = True
                Return boolIsParent
            Else
                boolIsParent = IsParentLineage(selBudgetItem, ParentBudgetItem)
                Return boolIsParent
            End If
        Else
            Return boolIsParent
        End If
    End Function

    Public Function GetRootActivity(ByVal ChildActivity As Activity) As Activity
        Dim ParentStruct As Struct = GetParent(ChildActivity)

        If ParentStruct.GetType() Is GetType(Activity) Then
            Return GetRootActivity(ParentStruct)
        Else
            Return ChildActivity
        End If
    End Function
#End Region

#Region "Get lists"
    Public Function GetOutputsList(Optional ByVal ParentPurposeGuid As Guid = Nothing) As Dictionary(Of Guid, String)
        Dim objOutputs As New Dictionary(Of Guid, String)

        objOutputs.Add(Guid.Empty, LANG_NotSelected)
        If ParentPurposeGuid = Nothing Or ParentPurposeGuid = Guid.Empty Then
            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    objOutputs.Add(selOutput.Guid, selOutput.Text)
                Next
            Next
        Else
            Dim selPurpose As Purpose = GetPurposeByGuid(ParentPurposeGuid)

            For Each selOutput As Output In selPurpose.Outputs
                objOutputs.Add(selOutput.Guid, selOutput.Text)
            Next
        End If

        Return objOutputs
    End Function

    Public Function GetActivitiesList(Optional ByVal ParentOutputGuid As Guid = Nothing) As Dictionary(Of Guid, String)
        Dim objActivitiesList As New Dictionary(Of Guid, String)

        objActivitiesList.Add(Guid.Empty, LANG_NotSelected)
        If ParentOutputGuid = Nothing Or ParentOutputGuid = Guid.Empty Then
            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    objActivitiesList = GetActivitiesList_Activities(selOutput.Activities, objActivitiesList)
                Next
            Next
        Else
            Dim selOutput As Output = GetOutputByGuid(ParentOutputGuid)

            objActivitiesList = GetActivitiesList_Activities(selOutput.Activities, objActivitiesList)
        End If

        Return objActivitiesList
    End Function

    Public Function GetActivitiesList_Activities(ByVal objActivities As Activities, ByVal objActivitiesList As Dictionary(Of Guid, String)) As Dictionary(Of Guid, String)
        For Each selActivity As Activity In objActivities
            objActivitiesList.Add(selActivity.Guid, selActivity.Text)

            If selActivity.Activities.Count > 0 Then
                objActivitiesList = GetActivitiesList_Activities(selActivity.Activities, objActivitiesList)
            End If
        Next

        Return objActivitiesList
    End Function

    Public Function GetResourcesList(Optional ByVal ParentActivityGuid As Guid = Nothing) As Dictionary(Of Guid, String)
        Dim objResourcesList As New Dictionary(Of Guid, String)

        objResourcesList.Add(Guid.Empty, LANG_NotSelected)
        If ParentActivityGuid = Nothing Or ParentActivityGuid = Guid.Empty Then
            For Each selPurpose As Purpose In Me.Purposes
                For Each selOutput As Output In selPurpose.Outputs
                    objResourcesList = GetResourcesList_Activities(selOutput.Activities, objResourcesList)
                Next
            Next
        Else
            Dim selActivity As Activity = GetActivityByGuid(ParentActivityGuid)

            For Each selResource As Resource In selActivity.Resources
                objResourcesList.Add(selResource.Guid, selResource.Text)
            Next
            If selActivity.Activities.Count > 0 Then
                objResourcesList = GetResourcesList_Activities(selActivity.Activities, objResourcesList)
            End If
        End If

        Return objResourcesList
    End Function

    Public Function GetResourcesList_Activities(ByVal objActivities As Activities, ByVal objResourcesList As Dictionary(Of Guid, String)) As Dictionary(Of Guid, String)
        For Each selActivity As Activity In objActivities
            For Each selResource As Resource In selActivity.Resources
                objResourcesList.Add(selResource.Guid, selResource.Text)
            Next
            If selActivity.Activities.Count > 0 Then
                objResourcesList = GetResourcesList_Activities(selActivity.Activities, objResourcesList)
            End If
        Next

        Return objResourcesList
    End Function

    Public Sub GetCustomMeasureUnits()
        lstUnits.Clear()
        MeasureUnitsUser.Clear()

        Dim MeasureUnits As List(Of StructuredComboBoxItem) = LoadMeasureUnits()
        For Each selItem As StructuredComboBoxItem In MeasureUnits
            If selItem.IsHeader = False Then
                lstUnits.Add(selItem.Unit)
            End If
        Next
        For Each selGoal As Goal In Me.Goals
            GetCustomMeasureUnits_Indicators(selGoal.Indicators)
        Next
        For Each selPurpose As Purpose In Me.Purposes
            GetCustomMeasureUnits_Indicators(selPurpose.Indicators)
            For Each selOutput As Output In selPurpose.Outputs
                GetCustomMeasureUnits_Indicators(selOutput.Indicators)
                GetCustomMeasureUnits_Activities(selOutput.Activities)
            Next
        Next
        LoadMeasureUnits()
    End Sub

    Private Sub GetCustomMeasureUnits_Activities(ByVal selActivities As Activities)
        For Each selActivity As Activity In selActivities
            GetCustomMeasureUnits_Indicators(selActivity.Indicators)

            If selActivity.Activities.Count > 0 Then GetCustomMeasureUnits_Activities(selActivity.Activities)
        Next
    End Sub

    Private Sub GetCustomMeasureUnits_Indicators(ByVal selIndicators As Indicators)
        For Each selInd As Indicator In selIndicators

            Select Case selInd.QuestionType
                Case Indicator.QuestionTypes.AbsoluteValue, Indicator.QuestionTypes.PercentageValue, _
                    Indicator.QuestionTypes.Ratio, Indicator.QuestionTypes.Formula

                    If selInd.ValuesDetail IsNot Nothing Then
                        With selInd.ValuesDetail
                            If lstUnits.Contains(.Unit) = False Then
                                MeasureUnitsUser.Add(New StructuredComboBoxItem(.Unit, False, False, , .Unit))
                                lstUnits.Add(.Unit)
                            End If
                        End With
                    End If

                    For Each selStatement As Statement In selInd.Statements
                        If selStatement.ValuesDetail IsNot Nothing Then
                            With selStatement.ValuesDetail
                                If lstUnits.Contains(.Unit) = False Then
                                    MeasureUnitsUser.Add(New StructuredComboBoxItem(.Unit, False, False, , .Unit))
                                    lstUnits.Add(.Unit)
                                End If
                            End With
                        End If
                    Next
            End Select
            If selInd.Indicators.Count > 0 Then
                GetCustomMeasureUnits_Indicators(selInd.Indicators)
            End If
        Next
    End Sub

    Public Function GetReferenceMomentsList(Optional ByVal objCurrentActivity As Activity = Nothing) As List(Of IdValuePair)
        Dim lstActivities As New List(Of IdValuePair)

        Dim Language As String = My.Settings.setLanguage

        lstActivities.Add(New IdValuePair(New Guid, LANG_NotSelected))
        For Each selMoment As KeyMoment In Me.KeyMomentsList
            lstActivities.Add(New IdValuePair(selMoment.Guid, selMoment.Description))
        Next

        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                lstActivities = GetReferenceMomenstsList_Activities(lstActivities, selOutput.Activities, objCurrentActivity)
            Next
        Next
        Return lstActivities
    End Function

    Private Function GetReferenceMomenstsList_Activities(ByVal lstActivities As List(Of IdValuePair), ByVal selActivities As Activities, ByVal objCurrentActivity As Activity) As List(Of IdValuePair)
        Dim strSortNumber As String

        For Each selActivity As Activity In selActivities
            If selActivity IsNot objCurrentActivity And IsParentLineage(objCurrentActivity, selActivity) = False Then
                strSortNumber = Me.GetStructSortNumber(selActivity)
                lstActivities.Add(New IdValuePair(selActivity.Guid, strSortNumber & " " & selActivity.Text))

            End If
            If selActivity.Activities.Count > 0 Then
                lstActivities = GetReferenceMomenstsList_Activities(lstActivities, selActivity.Activities, objCurrentActivity)
            End If
        Next

        Return lstActivities
    End Function

    Public Function GetTargetGroupsList() As Dictionary(Of Guid, String)
        Dim ListTargetGroups As New Dictionary(Of Guid, String)
        Dim strTargetGroup As String
        Dim intPurposeIndex, intIndex As Integer

        ListTargetGroups.Add(Guid.Empty, LANG_AllTargetGroups)
        For Each selPurpose As Purpose In Me.Purposes
            intPurposeIndex = Me.Purposes.IndexOf(selPurpose) + 1
            For Each selTargetGroup As TargetGroup In selPurpose.TargetGroups
                intIndex = selPurpose.TargetGroups.IndexOf(selTargetGroup) + 1
                strTargetGroup = String.Format("{0}.{1} {2}", intPurposeIndex, intIndex, selTargetGroup.Name)
                ListTargetGroups.Add(selTargetGroup.Guid, strTargetGroup)
            Next
        Next

        Return ListTargetGroups
    End Function

    Public Function GetTargetDeadlinesSection(ByVal intSection As Integer) As TargetDeadlinesSection
        Select Case intSection
            Case LogframeObject.SectionTypes.Goal
                Return Me.TargetDeadlinesGoals
            Case LogframeObject.SectionTypes.Purpose
                Return Me.TargetDeadlinesPurposes
            Case LogframeObject.SectionTypes.Output
                Return Me.TargetDeadlinesOutputs
            Case LogframeObject.SectionTypes.Activity
                Return Me.TargetDeadlinesActivities
        End Select

        Return Nothing
    End Function

    Public Function GetTargetDeadlines(ByVal intSection As Integer) As TargetDeadlines
        Select Case intSection
            Case LogframeObject.SectionTypes.Goal
                Return Me.TargetDeadlinesGoals.TargetDeadlines
            Case LogframeObject.SectionTypes.Purpose
                Return Me.TargetDeadlinesPurposes.TargetDeadlines
            Case LogframeObject.SectionTypes.Output
                Return Me.TargetDeadlinesOutputs.TargetDeadlines
            Case LogframeObject.SectionTypes.Activity
                Return Me.TargetDeadlinesActivities.TargetDeadlines
        End Select

        Return Nothing
    End Function

    Public Function GetDurationUnitsList() As List(Of String)
        Dim lstUnits As New List(Of String)
        Dim Language As String = My.Settings.setLanguage

        lstUnits.Add(String.Empty)
        lstUnits.Add(LANG_Days)
        lstUnits.Add(LANG_Weeks)
        lstUnits.Add(LANG_Months)
        lstUnits.Add(LANG_Years)

        For Each selYear As BudgetYear In Me.Budget.BudgetYears
            lstUnits = GetDurationUnitsList_BudgetItems(selYear.BudgetItems, lstUnits)
        Next

        Return lstUnits
    End Function

    Public Function GetDurationUnitsList_BudgetItems(ByVal objBudgetItems As BudgetItems, ByVal lstUnits As List(Of String)) As List(Of String)
        For Each selBudgetItem As BudgetItem In objBudgetItems
            If String.IsNullOrEmpty(selBudgetItem.DurationUnit) = False Then
                If lstUnits.Contains(selBudgetItem.DurationUnit) = False Then _
                    lstUnits.Add(selBudgetItem.DurationUnit)
            End If

            If selBudgetItem.BudgetItems.Count > 0 Then
                lstUnits = GetDurationUnitsList_BudgetItems(selBudgetItem.BudgetItems, lstUnits)
            End If
        Next

        Return lstUnits
    End Function

    Public Function GetNumberUnitsList() As List(Of StructuredComboBoxItem)
        Dim lstUnits As New List(Of StructuredComboBoxItem)

        lstUnits.Add(New StructuredComboBoxItem(LANG_Nothing, False, True, , String.Empty))
        lstUnits.Add(New StructuredComboBoxItem(LANG_FTE, False, False, , LANG_FTE))
        lstUnits.Add(New StructuredComboBoxItem(LANG_Items, False, False, , LANG_Items))
        lstUnits.Add(New StructuredComboBoxItem(LANG_Pieces, False, False, , LANG_Pieces))
        lstUnits.Add(New StructuredComboBoxItem(LANG_Persons, False, False, , LANG_Persons))
        lstUnits.Add(New StructuredComboBoxItem(LANG_Units, False, False, , LANG_Units))

        lstUnits.Add(New StructuredComboBoxItem(LANG_UserDefined, True, False))

        For Each selYear As BudgetYear In Me.Budget.BudgetYears
            lstUnits = GetNumberUnitsList_BudgetItems(selYear.BudgetItems, lstUnits)
        Next

        Return lstUnits
    End Function

    Public Function GetNumberUnitsList_BudgetItems(ByVal objBudgetItems As BudgetItems, ByVal lstUnits As List(Of StructuredComboBoxItem)) As List(Of StructuredComboBoxItem)
        For Each selBudgetItem As BudgetItem In objBudgetItems
            If String.IsNullOrEmpty(selBudgetItem.NumberUnit) = False Then
                SearchUnit = selBudgetItem.NumberUnit
                If lstUnits.Find(AddressOf FindExistingUnit) Is Nothing Then
                    Dim CustomUnit As New StructuredComboBoxItem(selBudgetItem.NumberUnit, False, False, , selBudgetItem.NumberUnit)
                    lstUnits.Add(CustomUnit)
                End If
            End If

            If selBudgetItem.BudgetItems.Count > 0 Then
                lstUnits = GetNumberUnitsList_BudgetItems(selBudgetItem.BudgetItems, lstUnits)
            End If
        Next

        Return lstUnits
    End Function

    Private Function FindExistingUnit(ByVal selItem As StructuredComboBoxItem) As Boolean
        If selItem.Description = SearchUnit Then Return True Else Return False
    End Function
#End Region

#Region "Get key moments"
    Public Function GetProjectStartKeyMoment() As KeyMoment
        For Each selMoment As KeyMoment In Me.Keymoments
            If selMoment.Type = KeyMoment.Types.ProjectStart Then
                Return selMoment
            End If
        Next

        Return Nothing
    End Function

    Public Function GetProjectEndKeyMoment() As KeyMoment
        For Each selMoment As KeyMoment In Me.Keymoments
            If selMoment.Type = KeyMoment.Types.ProjectEnd Then
                Return selMoment
            End If
        Next

        Return Nothing
    End Function

    Public Function GetFirstEvent(Optional ByVal boolIncludePreparation As Boolean = False) As Date
        Dim selDate As Date = Me.StartDate
        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                For Each selActivity As Activity In selOutput.Activities
                    With selActivity
                        If boolIncludePreparation = False Then
                            If .ExactStartDate > Date.MinValue And .ExactStartDate < selDate Then selDate = .ExactStartDate
                        Else
                            If .ActivityDetail.StartDatePreparation > Date.MinValue And .ActivityDetail.StartDatePreparation < selDate Then selDate = .ActivityDetail.StartDatePreparation
                        End If

                    End With
                Next
            Next
        Next
        selDate = DateSerial(selDate.Year, selDate.Month, selDate.Day)
        Return selDate
    End Function

    Public Function GetLastEvent(Optional ByVal boolIncludeFollowUp As Boolean = False) As Date
        Dim selDate As Date = Me.EndDate
        For Each selPurpose As Purpose In Me.Purposes
            For Each selOutput As Output In selPurpose.Outputs
                For Each selActivity As Activity In selOutput.Activities
                    With selActivity
                        If boolIncludeFollowUp = False Then
                            If .ExactEndDate > selDate Then selDate = .ExactEndDate
                        Else
                            If .ActivityDetail.EndDateFollowUp > Date.MinValue And .ActivityDetail.EndDateFollowUp > selDate Then selDate = .ActivityDetail.EndDateFollowUp
                        End If

                    End With

                Next
            Next
        Next
        selDate = DateSerial(selDate.Year, selDate.Month, selDate.Day)
        Return selDate
    End Function

    Public Function GetExecutionStart(Optional ByVal boolIncludePreparation As Boolean = False) As Date
        Dim datPeriodStart As Date
        If Me.StartDate > Date.MinValue Then _
            datPeriodStart = Me.StartDate
        If Me.GetFirstEvent > Date.MinValue And datPeriodStart > Me.GetFirstEvent Then
            datPeriodStart = Me.GetFirstEvent(boolIncludePreparation)
        End If
        datPeriodStart = DateSerial(datPeriodStart.Year, datPeriodStart.Month, datPeriodStart.Day)
        Return datPeriodStart
    End Function

    Public Function GetExecutionEnd(Optional ByVal boolIncludeFollowUp As Boolean = False) As Date
        Dim datPeriodEnd As Date
        If Me.EndDate > Date.MinValue Then _
        datPeriodEnd = Me.EndDate
        If Me.GetLastEvent > Date.MinValue And Me.EndDate < Me.GetLastEvent Then
            datPeriodEnd = Me.GetLastEvent(boolIncludeFollowUp)
        End If
        datPeriodEnd = DateSerial(datPeriodEnd.Year, datPeriodEnd.Month, datPeriodEnd.Day)
        Return datPeriodEnd
    End Function
#End Region

End Class
