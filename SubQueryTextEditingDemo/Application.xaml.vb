''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2023 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports GeneralAssembly.Windows

''' <summary>
''' Interaction logic for Application.xaml
''' </summary>
Partial Public Class Application
    Private Sub Application_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow As ExceptionWindow = New ExceptionWindow With {
                .Owner = Current.MainWindow,
                .Message = e.Exception.Message,
                .StackTrace = e.Exception.StackTrace
                }
        errorWindow.ShowDialog()
    End Sub
End Class
