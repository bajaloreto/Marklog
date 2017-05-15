Imports System.Reflection

Public Class ObjectCopy
    Implements System.IDisposable

    Public Function CopyObject(ByVal selObject As Object) As Object
        If selObject Is Nothing Then Return Nothing

        Dim objType As Type = selObject.GetType
        Dim objCopy As Object = Activator.CreateInstance(objType)
        Dim objProperties() As PropertyInfo = objType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)

        'copy values of properties
        For Each selProperty As PropertyInfo In objProperties
            If selProperty.CanWrite = True Then
                If selProperty.PropertyType.BaseType IsNot GetType(CollectionBase) And selProperty.PropertyType.BaseType IsNot GetType(Structs) Then
                    Dim selValue As Object

                    Select Case selProperty.PropertyType
                        Case GetType(Guid)
                            If selProperty.Name = "Guid" Then
                                selValue = Guid.NewGuid
                            Else
                                selValue = selProperty.GetValue(selObject, Nothing)
                            End If
                        Case GetType(ActivityDetail), GetType(AssumptionDetail), GetType(AudioVisualDetail), GetType(Baseline), _
                            GetType(Budget), GetType(Currency), GetType(DependencyDetail), GetType(OpenEndedDetail), GetType(RiskDetail), GetType(ScalesDetail), _
                            GetType(ValueRange), GetType(ValuesDetail)

                            selValue = selProperty.GetValue(selObject, Nothing)
                            selValue = CopyObject(selValue)
                        Case Else
                            selValue = selProperty.GetValue(selObject, Nothing)
                    End Select

                    selProperty.SetValue(objCopy, selValue, Nothing)
                Else
                    Dim selCollection As Object = selProperty.GetValue(selObject, Nothing)
                    Dim objCollectionType As Type = selCollection.GetType
                    Dim objCopyCollection As Object = Activator.CreateInstance(objCollectionType)
                    objCopyCollection = CopyCollection(selCollection)

                    selProperty.SetValue(objCopy, objCopyCollection, Nothing)
                End If
            End If
        Next

        Return objCopy
    End Function

    Public Function CopyCollection(ByVal objCollection As Object) As Object
        Dim objType As Type = objCollection.GetType
        Dim objCopyCollection As Object = Activator.CreateInstance(objType)

        Dim objCountProperty As PropertyInfo = objCollection.GetType.GetProperty("Count")
        If objCountProperty IsNot Nothing Then
            Dim intCount As Integer = objCountProperty.GetValue(objCollection, Nothing)
            If intCount > 0 Then
                For Each selChildObject As Object In objCollection
                    Dim objCopyChild As Object = CopyObject(selChildObject)
                    If objCopyChild IsNot Nothing Then
                        Dim objMethods As MethodInfo() = objType.GetMethods
                        'Dim objAddMethod As MethodInfo = objCollection.GetType.GetMethod("Add")

                        Dim objAddMethod As MethodInfo = objMethods.Single(Function(p) p.Name = "Add" AndAlso p.DeclaringType = objType)
                        'Dim propInfoSrcObj As PropertyInfo = myDE.[GetType]().GetProperties().[Single](Function(p) p.Name = "MyEntity" AndAlso p.PropertyType = GetType(MyDerivedEntity))
                        Dim objAttributes(0) As Object
                        objAttributes(0) = objCopyChild
                        If objAddMethod IsNot Nothing Then
                            objAddMethod.Invoke(objCopyCollection, objAttributes)
                        End If
                    End If
                Next
            End If
        End If
        Return objCopyCollection
    End Function

    Private disposedValue As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
