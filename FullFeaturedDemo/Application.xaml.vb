''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2022 Active Database Software              ''
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
    Public Shared Connections As New ConnectionList()
    Public Shared XmlFiles As New ConnectionList()

    Public Sub New()
        'if new version, import upgrade from previous version
        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
        End If

        If My.Settings.Connections IsNot Nothing Then
            Connections = My.Settings.Connections
        End If

        If My.Settings.XmlFiles IsNot Nothing Then
            XmlFiles = My.Settings.XmlFiles
        End If

        Connections.RestoreData()
        XmlFiles.RestoreData()
    End Sub

    Private Sub App_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow = New ExceptionWindow With {
            .Owner = Current.MainWindow,
            .Message = e.Exception.Message,
            .StackTrace = e.Exception.StackTrace
        }

        e.Handled = True

        errorWindow.ShowDialog()
    End Sub
End Class
