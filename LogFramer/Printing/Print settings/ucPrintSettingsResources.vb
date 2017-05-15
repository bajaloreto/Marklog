Imports System.Drawing.Printing

Public Class ucPrintSettingsResources
    Private intPagesWide As Integer
    Public boolLoad As Boolean

    Public Event PrintResourcesSetupChanged(ByVal sender As Object, ByVal e As PrintResourcesSetupChangedEventArgs)

    Public Sub New()
        InitializeComponent()

        boolLoad = True

        nudPagesWide.Value = My.Settings.setPrintResourcesPagesWide

        boolLoad = False
    End Sub

    Public Property PagesWide As Integer
        Get
            Return intPagesWide
        End Get
        Set(ByVal value As Integer)
            intPagesWide = value
        End Set
    End Property

    Private Sub nudPagesWide_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudPagesWide.ValueChanged
        Me.PagesWide = nudPagesWide.Value

        If boolLoad = False Then
            My.Settings.setPrintResourcesPagesWide = Me.PagesWide
            RaiseEvent PrintResourcesSetupChanged(Me, New PrintResourcesSetupChangedEventArgs(PagesWide))
        End If
    End Sub
End Class

Public Class PrintResourcesSetupChangedEventArgs
    Inherits EventArgs

    Public Property PagesWide As Integer

    Public Sub New(ByVal intPagesWide As Integer)
        MyBase.New()

        PagesWide = intPagesWide
    End Sub
End Class
