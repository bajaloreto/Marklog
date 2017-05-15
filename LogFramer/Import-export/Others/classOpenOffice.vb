Public Class classOpenOffice
    Public Sub Test()
        Dim objServiceManager As Object
        objServiceManager = CreateObject("com.sun.star.ServiceManager")
        Dim objDesktop As Object
        objDesktop = objServiceManager.createInstance("com.sun.star.frame.Desktop")
        Dim args As Object() = {}
        Dim objDocument As Object = objDesktop.LoadComponentFromURL("private:factory/scalc", "_blank", 0, args)

    End Sub
End Class
