Imports Microsoft.Win32

Public Class OfficeDetector
    ''' <summary>
    ''' Possible Office apps.
    ''' </summary>
    ''' <remarks></remarks>
    Enum MSOfficeApp
        Access_Application
        Excel_Application
        Outlook_Application
        PowerPoint_Application
        Word_Application
        FrontPage_Application
    End Enum

    ''' <summary>
    ''' Possible versions
    ''' </summary>
    ''' <remarks></remarks>
    Enum Version
        Version2007 = 12
        Version2003 = 11
        Version2002 = 10
        Version2000 = 9
        Version97 = 8
        Version95 = 7
    End Enum

    ''' <summary>
    ''' Lists all available MS Office apps and version in console window
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ListAllOfficeVersions()
        For Each s As String In [Enum].GetNames(GetType(MSOfficeApp))
            Console.WriteLine(s & vbTab & GetVersionsString([Enum].Parse(GetType(MSOfficeApp), s)))
        Next
    End Sub

    ''' <summary>
    ''' Returns version number as integer
    ''' </summary>
    ''' <param name="app"></param>
    ''' <returns></returns>
    ''' <remarks>Value is 0 if no version could be detected</remarks>
    Public Shared Function GetVersionsID(ByVal app As MSOfficeApp) As Integer
        Dim strProgID As String = [Enum].GetName(GetType(MSOfficeApp), app)
        strProgID = Replace(strProgID, "_", ".")
        Dim regKey As RegistryKey
        regKey = Registry.LocalMachine.OpenSubKey("Software\Classes\" & strProgID & "\CurVer", False)
        If IsNothing(regKey) Then Return 0
        Dim strV As String = regKey.GetValue("", Nothing, RegistryValueOptions.None)
        regKey.Close()

        strV = Replace(Replace(strV, strProgID, ""), ".", "")
        Return CInt(strV)
    End Function

    ''' <summary>
    ''' Returns the version string
    ''' </summary>
    ''' <param name="app"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetVersionsString(ByVal app As MSOfficeApp) As String
        Dim strProgID As String = [Enum].GetName(GetType(MSOfficeApp), app)
        strProgID = Replace(strProgID, "_", ".")
        Dim regKey As RegistryKey
        regKey = Registry.LocalMachine.OpenSubKey("Software\Classes\" & strProgID & "\CurVer", False)
        If IsNothing(regKey) Then Return "No version detected."
        Dim strV As String = regKey.GetValue("", Nothing, RegistryValueOptions.None)
        regKey.Close()

        strV = Replace(Replace(strV, strProgID, ""), ".", "")
        Return [Enum].GetName(GetType(Version), CInt(strV))
    End Function
End Class
