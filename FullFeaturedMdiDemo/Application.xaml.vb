''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports GeneralAssembly
Imports GeneralAssembly.Windows

''' <summary>
''' Interaction logic for App.xaml
''' </summary>
Partial Public Class App
    Inherits Application

    Public Shared Name As String = "Active Query Builder Demo"

    Public Shared Connections As New ConnectionList()
    Public Shared XmlFiles As New ConnectionList()

    Public Sub New()
        Dim i = ControlFactory.Instance ' force call static constructor of control factory

        'if new version, import upgrade from previous version
        If My.Settings.Default.CallUpgrade Then
            My.Settings.Default.Upgrade()
            My.Settings.Default.CallUpgrade = False
        End If

        If My.Settings.Default.Connections IsNot Nothing Then
            Connections = My.Settings.Default.Connections
            Connections.RemoveObsoleteConnectionInfos()
            Connections.RestoreData()
        End If

        If My.Settings.Default.XmlFiles IsNot Nothing Then
            XmlFiles = My.Settings.Default.XmlFiles
            XmlFiles.RemoveObsoleteConnectionInfos()
            XmlFiles.RestoreData()
        End If
    End Sub

    Private Sub App_OnExit(sender As Object, e As ExitEventArgs)
        Connections.SaveData()
        XmlFiles.SaveData()
        My.Settings.Default.Connections = Connections
        My.Settings.Default.XmlFiles = XmlFiles
        My.Settings.Default.Save()
    End Sub

    Private Sub App_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow = New ExceptionWindow With {
            .Owner = Current.MainWindow,
            .Message = e.Exception.Message,
            .StackTrace = e.Exception.StackTrace
        }

        errorWindow.ShowDialog()
    End Sub
End Class
