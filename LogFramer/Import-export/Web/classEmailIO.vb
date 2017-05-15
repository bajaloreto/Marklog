Imports System.Text

Public Class EmailIO
    Implements System.IDisposable

    Public Sub StartMailClient(ByVal selEmail As Email)
        'string builder used for concatination
        Dim MsgBuilder As New StringBuilder
        Dim strTo As String = String.Format("mailto:{0}", selEmail.Email)
        Dim strFileName As String

        MsgBuilder.Append(strTo)

        strFileName = MsgBuilder.ToString

        Using objInternet As New classInternet
            objInternet.ExecuteFile(strFileName, ERR_CannotStartEmailClient)
        End Using


        'System.Diagnostics.Process proc = new System.Diagnostics.Process();
        'proc.StartInfo.FileName = "mailto:someone@somewhere.com?subject=hello&body=love my body";
        'proc.Start();
    End Sub

    

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
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
