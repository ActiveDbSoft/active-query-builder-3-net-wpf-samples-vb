//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports GeneralAssembly.Windows

''' <summary>
''' Interaction logic for App.xaml
''' </summary>
Partial Public Class App
    Private Sub App_OnDispatcherUnhandledException(ByVal sender As Object, ByVal e As DispatcherUnhandledExceptionEventArgs)
        Dim errorWindow = New ExceptionWindow With {
            .Owner = Current.MainWindow,
            .Message = e.Exception.Message,
            .StackTrace = e.Exception.StackTrace
        }

        errorWindow.ShowDialog()
    End Sub
End Class
