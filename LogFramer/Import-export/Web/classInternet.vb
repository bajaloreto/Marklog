Public Class classInternet
    Implements System.IDisposable

    Public Function ExecuteFile(ByVal strFileName As String, ByVal strError As String) As Boolean
        If String.IsNullOrEmpty(strFileName) = False Then
            Try
                Using myProcess As New System.Diagnostics.Process
                    With myProcess

                        .StartInfo.FileName = strFileName
                        '.StartInfo.UseShellExecute = True
                        '.StartInfo.RedirectStandardOutput = False
                        .Start()
                    End With
                End Using
            Catch ex As System.ComponentModel.Win32Exception
                MsgBox(strError, MsgBoxStyle.OkOnly)
            End Try
        End If
    End Function

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
