'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2018 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes

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
