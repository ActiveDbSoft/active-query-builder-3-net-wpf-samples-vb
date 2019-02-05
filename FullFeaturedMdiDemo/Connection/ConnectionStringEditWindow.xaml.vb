'*******************************************************************'
'       Active Query Builder Component Suite                        '
'                                                                   '
'       Copyright © 2006-2019 Active Database Software              '
'       ALL RIGHTS RESERVED                                         '
'                                                                   '
'       CONSULT THE LICENSE AGREEMENT FOR INFORMATION ON            '
'       RESTRICTIONS.                                               '
'*******************************************************************'

Imports System.Windows

Namespace Connection
	''' <summary>
	''' Interaction logic for ConnectionStringEditWindow.xaml
	''' </summary>
	Public Partial Class ConnectionStringEditWindow
		Public Property ConnectionString() As String
			Get
				Return tbConnectionString.Text
			End Get
			Set
				tbConnectionString.Text = value
				Modified = False
			End Set
		End Property

		Public Property Modified() As Boolean
			Get
				Return m_Modified
			End Get
			Private Set
				m_Modified = Value
			End Set
		End Property
		Private m_Modified As Boolean

		Public Sub New()
			InitializeComponent()
			Modified = True
		End Sub

		Private Sub ButtonBaseOk_OnClick(sender As Object, e As RoutedEventArgs)
			DialogResult = True
			Close()
		End Sub

		Private Sub ButtonBaseClose_OnClick(sender As Object, e As RoutedEventArgs)
			Close()
		End Sub
	End Class
End Namespace
