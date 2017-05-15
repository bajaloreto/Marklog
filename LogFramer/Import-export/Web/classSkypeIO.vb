'Imports SKYPE4COMLib

Public Class SkypeIO
    Implements System.IDisposable

    Public Sub MakeCall(ByVal strSkypeAccount As String)
        'Dim objSkype As New Skype
        Dim strCall As String

        Try
            'If Not objSkype.Client.IsRunning Then
            '    objSkype.Client.Start()
            'End If
            'objSkype.Attach(objSkype.Protocol, True)
            'objSkype.PlaceCall(strSkypeAccount)

            strCall = String.Format("skype:{0}", strSkypeAccount)
            Using objInternet As New classInternet
                objInternet.ExecuteFile(strCall, ERR_CannotStartSkype)
            End Using

        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox("Not able to connect to Skype. Make sure Skype is installed on your computer.", MsgBoxStyle.OkOnly)
        End Try


    End Sub


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
