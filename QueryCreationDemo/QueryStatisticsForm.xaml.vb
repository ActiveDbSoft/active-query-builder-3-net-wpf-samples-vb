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
''' Interaction logic for QueryStatisticsForm.xaml
''' </summary>
Public Partial Class QueryStatisticsForm
	Inherits Window
	Public Sub New()
		InitializeComponent()
	End Sub

	Public Sub New(sql As String)
        SqlTextBox.Text = sql
	End Sub

	Private Sub ButtonOk_OnClick(sender As Object, e As RoutedEventArgs)
		DialogResult = False
	End Sub
End Class
