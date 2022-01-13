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
    Public Sub New()
        'if new version, import upgrade from previous version

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
