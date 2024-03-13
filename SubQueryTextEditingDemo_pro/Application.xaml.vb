''*******************************************************************''
''       Active Query Builder Component Suite                        ''
''                                                                   ''
''       Copyright © 2006-2024 Active Database Software              ''
''       ALL RIGHTS RESERVED                                         ''
''                                                                   ''
''       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            ''
''       RESTRICTIONS.                                               ''
''*******************************************************************''

Imports GeneralAssembly.Windows

Namespace SubQueryTextEditingDemo
    ''' <summary>
    ''' Interaction logic for App.xaml
    ''' </summary>
    Partial Public Class App
        Private Sub App_OnDispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs)
            Dim errorWindow = New ExceptionWindow With {
                .Owner = Current.MainWindow,
                .Message = e.Exception.Message,
                .StackTrace = e.Exception.StackTrace
            }

            errorWindow.ShowDialog()
        End Sub
    End Class
End Namespace
