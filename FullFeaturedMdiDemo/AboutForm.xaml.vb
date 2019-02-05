'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Diagnostics
Imports System.Reflection
Imports System.Windows.Navigation
Imports ActiveQueryBuilder.View.WPF

''' <summary>
''' Interaction logic for AboutForm.xaml
''' </summary>
Public Partial Class AboutForm
	Public Sub New()
		InitializeComponent()

		LblQueryBuilderVersion.Text = "v" & Convert.ToString(GetType(QueryBuilder).Assembly.GetName().Version)
		LblDemoVersion.Text = "v" & Convert.ToString(Assembly.GetExecutingAssembly().GetName().Version)
	End Sub

	Private Sub Hyperlink1_RequestNavigate(sender As Object, e As RequestNavigateEventArgs)
		Process.Start(New ProcessStartInfo(e.Uri.AbsoluteUri))
		e.Handled = True
	End Sub
End Class
