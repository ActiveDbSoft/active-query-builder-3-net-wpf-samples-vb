'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows.Threading
Imports Windows

''' <summary>
''' Interaction logic for App.xaml
''' </summary>
Partial Public Class App
    Public Shared Connections As New ConnectionList()
    Public Shared XmlFiles As New ConnectionList()

    Public Sub New()
        'if new version, import upgrade from previous version
        If My.Settings.Default.CallUpgrade Then
            My.Settings.Default.Upgrade()
            My.Settings.Default.CallUpgrade = False
        End If

        If My.Settings.Default.Connections IsNot Nothing Then
            Connections = My.Settings.Default.Connections
        End If

        If My.Settings.Default.XmlFiles IsNot Nothing Then
            XmlFiles = My.Settings.Default.XmlFiles
        End If

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
