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
Imports LoadMetadataDemo.Common

''' <summary>
''' Interaction logic for App.xaml
''' </summary>
Public Partial Class App
	Inherits Application
	Private Sub Application_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow As ExceptionWindow = New ExceptionWindow With {
                .Owner = Current.MainWindow,
                .Message = e.Exception.Message,
                .StackTrace = e.Exception.StackTrace
                }
        errorWindow.ShowDialog()
    End Sub
End Class
