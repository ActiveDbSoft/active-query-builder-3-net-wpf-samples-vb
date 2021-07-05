//*******************************************************************//
//       Active Query Builder Component Suite                        //
//                                                                   //
//       Copyright © 2006-2021 Active Database Software              //
//       ALL RIGHTS RESERVED                                         //
//                                                                   //
//       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            //
//       RESTRICTIONS.                                               //
//*******************************************************************//

Imports System.Windows

Namespace Windows.QueryInformationWindows
	''' <summary>
	''' Interaction logic for QueryStatisticsWindow.xaml
	''' </summary>
	Partial Public Class QueryStatisticsWindow
		Inherits Window

		Public Sub New(text As String)
			InitializeComponent()
			textBox.Text = text
		End Sub

		Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub
	End Class
End Namespace
