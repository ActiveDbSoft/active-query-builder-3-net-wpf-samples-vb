''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2021 Active Database Software              ''
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
        If Settings.Default.CallUpgrade Then
            Settings.Default.Upgrade()
            Settings.Default.CallUpgrade = False
        End If

        If Settings.Default.Connections IsNot Nothing Then
            Connections = Settings.Default.Connections
        End If

        If Settings.Default.XmlFiles IsNot Nothing Then
            XmlFiles = Settings.Default.XmlFiles
        End If

        Settings.Default.Connections = Connections
        Settings.Default.XmlFiles = XmlFiles
        Settings.Default.Save()
    End Sub

    Private Sub App_OnDispatcherUnhandledException(ByVal sender As Object, ByVal e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow = New ExceptionWindow With {
            .Owner = Current.MainWindow,
            .Message = e.Exception.Message,
            .StackTrace = e.Exception.StackTrace
        }

        errorWindow.ShowDialog()
    End Sub
End Class
