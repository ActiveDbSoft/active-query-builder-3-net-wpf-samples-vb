'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'


Imports System.Windows

''' <summary>
''' Interaction logic for QueryStatisticsWindow.xaml
''' </summary>
Public Partial Class QueryStatisticsWindow
	Inherits Window
	Public Sub New()
		InitializeComponent()
	End Sub

	Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
		Close()
	End Sub
End Class
