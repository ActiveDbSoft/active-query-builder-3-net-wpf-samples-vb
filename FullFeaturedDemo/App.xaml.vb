﻿'*******************************************************************'
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
Imports FullFeaturedDemo.Common
Imports FullFeaturedDemo.Properties

''' <summary>
''' Interaction logic for App.xaml
''' </summary>
Public Partial Class App
	Public Shared Connections As New ConnectionList()
	Public Shared XmlFiles As New ConnectionList()

	Public Sub New()
		'if new version, import upgrade from previous version
		If Settings.[Default].CallUpgrade Then
			Settings.[Default].Upgrade()
			Settings.[Default].CallUpgrade = False
		End If

		If Settings.[Default].Connections IsNot Nothing Then
			Connections = Settings.[Default].Connections
		End If

		If Settings.[Default].XmlFiles IsNot Nothing Then
			XmlFiles = Settings.[Default].XmlFiles
		End If

		Settings.[Default].Connections = Connections
		Settings.[Default].XmlFiles = XmlFiles
		Settings.[Default].Save()
	End Sub
	
	Private Sub Application_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow = New ExceptionWindow With {
                .Owner = Current.MainWindow,
                .Message = e.Exception.Message,
                .StackTrace = e.Exception.StackTrace
                }
        errorWindow.ShowDialog()
    End Sub
End Class
