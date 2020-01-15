'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'


Imports System.Windows
Imports System.Windows.Threading
Imports ActiveQueryBuilder.View.WPF
Imports FullFeaturedMdiDemo.Common
Imports FullFeaturedMdiDemo.Properties


    Public Partial Class App
        Inherits Application

        Public Shared Name As String = "Active Query Builder Demo"
        Public Shared Connections As ConnectionList = New ConnectionList()
        Public Shared XmlFiles As ConnectionList = New ConnectionList()

        Public Sub New()
            Dim i = ControlFactory.Instance

            If Settings.Default.CallUpgrade Then
                Settings.Default.Upgrade()
                Settings.Default.CallUpgrade = False
            End If

            If Settings.Default.Connections IsNot Nothing Then
                Connections = Settings.Default.Connections
                Connections.RemoveObsoleteConnectionInfos()
                Connections.RestoreData()
            End If

            If Settings.Default.XmlFiles IsNot Nothing Then
                XmlFiles = Settings.Default.XmlFiles
                XmlFiles.RemoveObsoleteConnectionInfos()
                XmlFiles.RestoreData()
            End If
        End Sub

        Private Sub App_OnExit(sender As Object, e As ExitEventArgs)
            Connections.SaveData()
            XmlFiles.SaveData()
            Settings.Default.Connections = Connections
            Settings.Default.XmlFiles = XmlFiles
            Settings.Default.Save()
        End Sub

        Private Sub App_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
            Dim errorWindow As ExceptionWindow = New ExceptionWindow With {
                .Owner = Current.MainWindow,
                .Message = e.Exception.Message,
                .StackTrace = e.Exception.StackTrace
            }
            errorWindow.ShowDialog()
        End Sub
    End Class
