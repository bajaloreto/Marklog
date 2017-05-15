'Imports Microsoft.Office
'Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class OutlookIO
    Public Sub CreateNewMessage(Optional ByVal strTo As String = "", Optional ByVal strCC As String = "", _
                                Optional ByVal strSubject As String = "", Optional ByVal strBody As String = "")
        Try
            ' Create the Outlook application by using inline initialization.
            Dim objOutlook As Object = CreateObject("Outlook.Application")

            'Create the new message by using the simplest approach.
            Dim objMessage As Object = objOutlook.CreateItem(0)

            'Set the basic properties.
            If strTo <> "" Then objMessage.To = strTo
            If strCC <> "" Then objMessage.CC = strCC
            If strSubject <> "" Then objMessage.Subject = strSubject
            If strBody <> "" Then objMessage.Body = strBody

            ' If you want to, display the message.
            objMessage.Display(True)


            'Explicitly release objects.
            objMessage = Nothing
            objOutlook = Nothing

        Catch e As Exception
            Try
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = "mailto:" & strTo & "?subject=" & strSubject & "&body=" & strBody
                proc.Start()
            Catch ex As Exception
                MsgBox(e.Message)
            End Try

        End Try
    End Sub
End Class
